import datetime
import operator

from loguru import logger
from sqlalchemy import Column, String, create_engine, Integer, Date
from sqlalchemy import and_
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
from tqdm import tqdm

import utils.data_util as data

engine = create_engine('mysql+pymysql://root:slmlms123@localhost:3306/jian_yan_pi', echo=False)
Base = declarative_base()
DBSession = sessionmaker(bind=engine)
table_name = "消防"
fenbu_num = "S8"
fenbu_name = "通讯分部工程质量验收记录"
session = DBSession()
# 日期格式
date_formate = "%Y.%m.%d"
zifenbus_dq = ["室外电气", "变配电室", "供电干线", "电气动力", "电气照明", "备用和不间断电源", "防雷及接地", "仪表安装", "管线敷设", "控制系统接地", "通信网络系统",
               "计算机网络系统",
               "工业监控系统", "火灾报警系统"]
zifenbus_yb = ["仪表安装", "管线敷设", "控制系统接地"]
zifenbus_dx = ["通信网络系统", "计算机网络系统", "工业监控系统"]
zifenbus_xf = ["火灾报警系统"]

zixiangs = ["粗矿堆", "粗碎站", "浸出浓密及CCD洗涤", "磨矿分级", "皮带廊及转运站", "铜萃取", "铜电积", "原矿脱水及搅拌浸出", "总降压变电站", "给水加压泵房及回水池", "综合管网",
            "新水泵房及输送管线", "硫酸库", "溶剂油库", "选厂设备循环水站", "空压机站", "溶液精滤", "钴回收", "石灰乳及石灰石浆制备"]


# zixiangs = ["粗矿堆"]


class fenXiang(Base):
    __tablename__ = '分项工程-' + table_name
    id = Column(Integer, primary_key=True)
    子项名称 = Column(String(20))
    分项工程代号 = Column(Integer)
    分项工程名称 = Column(String(20))
    检验批数量 = Column(Integer)
    最后一个检验批时间 = Column(Date)


class jian_yan_pi(Base):
    __tablename__ = table_name
    id = Column(Integer, primary_key=True)
    子项名称 = Column(String(20))
    分项工程名称 = Column(String(20))
    分项工程代号 = Column(Integer)
    真报验时间 = Column(Date)
    真编号 = Column(String(20))


class mu_ban(Base):
    __tablename__ = '模板'
    id = Column(Integer, primary_key=True)
    分项工程名称 = Column(String(20))
    子分部工程名称 = Column(String(20))
    分部工程代号 = Column(String(20))
    子分部代号 = Column(Integer)


class yin_bi(Base):
    __tablename__ = '隐蔽工程'
    id = Column(Integer, primary_key=True)
    子项名称 = Column(String(20))
    隐蔽类型 = Column(String(20))
    报验时间 = Column(Date)


class zhao_ming(Base):
    __tablename__ = '建筑照明试运行记录'
    id = Column(Integer, primary_key=True)
    子项名称 = Column(String(20))
    开始时间 = Column(Date)


class dian_ji(Base):
    __tablename__ = '电机实验记录'
    id = Column(Integer, primary_key=True)
    子项名称 = Column(String(20))
    检查日期 = Column(String(20))


class bian_pin_qi(Base):
    __tablename__ = '变频器实验记录'
    id = Column(Integer, primary_key=True)
    子项名称 = Column(String(20))
    检查日期 = Column(String(20))


class ruan_qi(Base):
    __tablename__ = '软启实验记录'
    id = Column(Integer, primary_key=True)
    子项名称 = Column(String(20))
    检查日期 = Column(String(20))


class mu_xian(Base):
    __tablename__ = '母线实验记录'
    id = Column(Integer, primary_key=True)
    子项名称 = Column(String(20))
    真检查日期 = Column(String(20))


class duan_lu_qi(Base):
    __tablename__ = '10kv断路器'
    id = Column(Integer, primary_key=True)
    安装地点 = Column(String(20))
    实验日期 = Column(String(20))


class song_pei_dian(Base):
    __tablename__ = '10kv配电系统'
    id = Column(Integer, primary_key=True)
    子项名称 = Column(String(20))
    检查日期 = Column(String(20))


class bian_ya_qi(Base):
    __tablename__ = '变压器实验记录'
    id = Column(Integer, primary_key=True)
    子项名称 = Column(String(20))
    位号 = Column(String(20))
    实验日期 = Column(Date)


class jie_di_dian_zu(Base):
    __tablename__ = '接地电阻测试记录'
    id = Column(Integer, primary_key=True)
    子项名称 = Column(String(20))
    编号 = Column(String(20))
    部位 = Column(String(20))
    实验日期 = Column(Date)


class fen_bu_zi_fen_bu(Base):
    __tablename__ = '分部子分部'
    id = Column(Integer, primary_key=True)
    子项名称 = Column(String(20))
    分部名称 = Column(String(20))
    日期 = Column(Date)


@logger.catch()
def export_menu(row_num, zifenbus):
    # Excel模板，注意选择打开方式
    excel_template = data.switch_open_excel_template("openpyxl", "inspection_lot/卷内目录 .xlsx")

    # 子项名称
    excel_template.worksheets[0]["A2"].value = "子项名称：" + s
    # 分部
    if operator.eq(table_name, "消防") == False:
        excel_template.worksheets[0]["B" + str(row_num)].value = "--"
        excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
        excel_template.worksheets[0]["D" + str(row_num)].value = fenbu_name
        if session.query(fen_bu_zi_fen_bu).filter(
                and_(fen_bu_zi_fen_bu.子项名称 == s, fen_bu_zi_fen_bu.分部名称 == fenbu_name)).count() > 0:
            fen_bu_date = session.query(fen_bu_zi_fen_bu).filter(
                and_(fen_bu_zi_fen_bu.子项名称 == s, fen_bu_zi_fen_bu.分部名称 == fenbu_name)).first().日期
            excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(fen_bu_date,
                                                                                                date_formate)
        row_num = row_num + 1
    for zifenbu in zifenbus:
        list_mb = session.query((mu_ban)).filter(and_(mu_ban.子分部工程名称 == zifenbu, mu_ban.分部工程代号 == fenbu_num)).all()
        if len(list_mb) < 1: continue
        count = session.query(fenXiang).filter(
            and_(fenXiang.子项名称 == s,
                 fenXiang.检验批数量 > 0, fenXiang.分项工程代号 == list_mb[0].子分部代号)).count()
        # 子分部
        if count == 0: continue
        excel_template.worksheets[0]["B" + str(row_num)].value = "--"
        excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
        excel_template.worksheets[0]["D" + str(row_num)].value = zifenbu + "子分部工程质量验收记录"
        if session.query(fen_bu_zi_fen_bu).filter(
                and_(fen_bu_zi_fen_bu.子项名称 == s, fen_bu_zi_fen_bu.分部名称 == zifenbu + "子分部工程质量验收记录")).count() > 0:
            zi_fen_bu_date = session.query(fen_bu_zi_fen_bu).filter(
                and_(fen_bu_zi_fen_bu.子项名称 == s, fen_bu_zi_fen_bu.分部名称 == zifenbu + "子分部工程质量验收记录")).first().日期
            excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(zi_fen_bu_date,
                                                                                                date_formate)
        row_num = row_num + 1
        for lm in list_mb:
            list = session.query(fenXiang).filter(
                and_(fenXiang.分项工程名称 == lm.分项工程名称, fenXiang.子项名称 == s,
                     fenXiang.检验批数量 > 0, fenXiang.分项工程代号 == lm.子分部代号)).order_by(
                fenXiang.分项工程代号).all()
            if len(list) == 0:
                continue
            for l in list:
                # 分项工程
                excel_template.worksheets[0]["B" + str(row_num)].value = "--"
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                excel_template.worksheets[0]["D" + str(row_num)].value = l.分项工程名称 + "分项工程质量验收记录"
                excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                    datetime.datetime.strptime(l.最后一个检验批时间, '%Y-%m-%d'), date_formate)
                row_num = row_num + 1

                jyp = session.query(jian_yan_pi).filter(and_(jian_yan_pi.子项名称 == s, jian_yan_pi.分项工程代号 == l.分项工程代号,
                                                             jian_yan_pi.真编号.like(
                                                                 "%" + fenbu_num + "-" + l.分项工程代号.zfill(
                                                                     2) + "-C7-%"),
                                                             jian_yan_pi.分项工程名称 == l.分项工程名称)).order_by(
                    jian_yan_pi.真报验时间).all()
                # 编号
                i = int(l.检验批数量)
                # 设置起止编号和起止日期
                if i == 1:
                    bianhao = fenbu_num + "-" + l.分项工程代号.zfill(2) + "-C7-" + "{:0>3d}".format(1)
                    jyp_date = datetime.datetime.strftime(
                        datetime.datetime.strptime(jyp[0].真报验时间, '%Y-%m-%d'), date_formate)
                else:
                    excel_template.worksheets[0]["F" + str(row_num)].value = i
                    bianhao = fenbu_num + "-" + l.分项工程代号.zfill(2) + "-C7-" + "{:0>3d}".format(
                        1) + "~" + "{:0>3d}".format(i)
                    if jyp[0].真报验时间 == jyp[len(jyp) - 1].真报验时间:
                        jyp_date = datetime.datetime.strftime(
                            datetime.datetime.strptime(jyp[0].真报验时间, '%Y-%m-%d'), date_formate)
                    else:
                        jyp_date = datetime.datetime.strftime(
                            datetime.datetime.strptime(jyp[0].真报验时间, '%Y-%m-%d'),
                            date_formate) + "~" + datetime.datetime.strftime(
                            datetime.datetime.strptime(jyp[len(jyp) - 1].真报验时间, '%Y-%m-%d'), date_formate)

                # assert jyp is not None
                # 检验批写入
                excel_template.worksheets[0]["B" + str(row_num)].value = bianhao
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                excel_template.worksheets[0]["D" + str(row_num)].value = l.分项工程名称 + "检验批质量验收记录"
                excel_template.worksheets[0]["E" + str(row_num)].value = jyp_date

                row_num = row_num + 1

        if (zifenbu == "变配电室"):
            list_mx = session.query(mu_xian).filter(mu_xian.子项名称 == s).order_by(mu_xian.真检查日期).all()
            if len(list_mx) > 0:
                excel_template.worksheets[0]["D" + str(row_num)].value = "母线实验记录"
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                if len(list_mx) == 1 or list_mx[0].真检查日期 == list_mx[len(list_mx) - 1].真检查日期:
                    excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                        datetime.datetime.strptime(list_mx[0].真检查日期, '%Y-%m-%d'), date_formate)
                else:
                    excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                        datetime.datetime.strptime(list_mx[0].真检查日期, '%Y-%m-%d'),
                        date_formate) + "~" + datetime.datetime.strftime(
                        datetime.datetime.strptime(list_mx[len(list_mx) - 1].真检查日期, '%Y-%m-%d'), date_formate)
                    excel_template.worksheets[0]["F" + str(row_num)].value = len(list_mx)
                excel_template.worksheets[0]["B" + str(row_num)].value = "--"
                row_num = row_num + 1
            list_byq = session.query(bian_ya_qi).filter(bian_ya_qi.子项名称 == s).order_by(bian_ya_qi.实验日期).all()
            for byq in list_byq:
                excel_template.worksheets[0]["D" + str(row_num)].value = "变压器实验记录(" + byq.位号 + ")"
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                excel_template.worksheets[0]["B" + str(row_num)].value = "--"
                excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(byq.实验日期,
                                                                                                    date_formate)
                row_num = row_num + 1

            if (s == "总降压变电站" or s == "磨矿分级" or s == "钴回收"):
                if s == "总降压变电站":
                    excel_template.worksheets[0]["D" + str(row_num)].value = "GIS实验记录"
                    excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                    excel_template.worksheets[0]["E" + str(row_num)].value = "2020.05.25"
                    excel_template.worksheets[0]["B" + str(row_num)].value = "--"

                    excel_template.worksheets[0]["D" + str(row_num + 1)].value = "1#主变压器实验记录"
                    excel_template.worksheets[0]["C" + str(row_num + 1)].value = "北方国际"
                    excel_template.worksheets[0]["E" + str(row_num + 1)].value = "2020.05.25"
                    excel_template.worksheets[0]["B" + str(row_num + 1)].value = "--"

                    excel_template.worksheets[0]["D" + str(row_num + 2)].value = "2#主变压器实验记录"
                    excel_template.worksheets[0]["C" + str(row_num + 2)].value = "北方国际"
                    excel_template.worksheets[0]["E" + str(row_num + 2)].value = "2020.05.25"
                    excel_template.worksheets[0]["B" + str(row_num + 2)].value = "--"
                    row_num = row_num + 3
                list_dlq = session.query(duan_lu_qi).filter(duan_lu_qi.安装地点 == s + "10kV配电室").order_by(
                    duan_lu_qi.实验日期).all()
                if len(list_dlq) > 0:
                    excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                    excel_template.worksheets[0]["D" + str(row_num)].value = "10kV断路器实验记录"
                    excel_template.worksheets[0]["B" + str(row_num)].value = "--"
                    excel_template.worksheets[0]["F" + str(row_num)].value = len(list_dlq)
                    if len(list_dlq) == 1 or list_dlq[0].实验日期 == list_dlq[len(list_dlq) - 1].实验日期:
                        excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                            datetime.datetime.strptime(list_dlq[0].实验日期, '%Y年%m月%d日'),
                            date_formate)
                    else:
                        excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                            datetime.datetime.strptime(list_dlq[0].实验日期, '%Y年%m月%d日'),
                            date_formate) + "~" + datetime.datetime.strftime(
                            datetime.datetime.strptime(list_dlq[len(list_dlq) - 1].实验日期, '%Y年%m月%d日'), date_formate)

                    row_num = row_num + 1
                list_spd = session.query(song_pei_dian).filter(song_pei_dian.子项名称 == s).order_by(
                    song_pei_dian.检查日期).all()
                if len(list_spd) > 0:
                    excel_template.worksheets[0]["D" + str(row_num)].value = "10kV送配电系统调试记录"
                    excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                    excel_template.worksheets[0]["B" + str(row_num)].value = "--"
                    excel_template.worksheets[0]["F" + str(row_num)].value = len(list_spd)
                    if len(list_spd) == 1 or list_spd[0].检查日期 == list_spd[len(list_spd) - 1].检查日期:
                        excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                            datetime.datetime.strptime(list_spd[0].检查日期, '%Y-%m-%d'),
                            date_formate)
                    else:
                        excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                            datetime.datetime.strptime(list_spd[0].检查日期, '%Y-%m-%d'),
                            date_formate) + "~" + datetime.datetime.strftime(
                            datetime.datetime.strptime(list_spd[len(list_spd) - 1].检查日期, '%Y-%m-%d'), date_formate)

                    row_num = row_num + 1

                excel_template.worksheets[0]["D" + str(row_num)].value = "过电压保护器调试记录"
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                excel_template.worksheets[0]["B" + str(row_num)].value = "--"
                if s == "总降压变电站":
                    excel_template.worksheets[0]["E" + str(row_num)].value = "2020.05.06"
                elif s == "磨矿分级":
                    excel_template.worksheets[0]["E" + str(row_num)].value = "2020.05.25"
                elif s == "钴回收":
                    excel_template.worksheets[0]["E" + str(row_num)].value = "2020.11.15"
                row_num = row_num + 1

                excel_template.worksheets[0]["D" + str(row_num)].value = "电流互感器实验记录"
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                excel_template.worksheets[0]["B" + str(row_num)].value = "--"
                if s == "总降压变电站":
                    excel_template.worksheets[0]["E" + str(row_num)].value = "2020.04.28"
                elif s == "磨矿分级":
                    excel_template.worksheets[0]["E" + str(row_num)].value = "2020.06.17"
                elif s == "钴回收":
                    excel_template.worksheets[0]["E" + str(row_num)].value = "2020.11.11"
                row_num = row_num + 1

                excel_template.worksheets[0]["D" + str(row_num)].value = "继电保护装置调试记录"
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                excel_template.worksheets[0]["B" + str(row_num)].value = "--"
                if s == "总降压变电站":
                    excel_template.worksheets[0]["E" + str(row_num)].value = "2020.07.03~2020.07.19"
                elif s == "磨矿分级":
                    excel_template.worksheets[0]["E" + str(row_num)].value = "2020.06.17~2020.07.16"
                elif s == "钴回收":
                    excel_template.worksheets[0]["E" + str(row_num)].value = "2020.12.25"
                row_num = row_num + 1

                excel_template.worksheets[0]["D" + str(row_num)].value = "高压电缆实验记录"
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                excel_template.worksheets[0]["B" + str(row_num)].value = "--"
                if s == "总降压变电站":
                    excel_template.worksheets[0]["E" + str(row_num)].value = "2020.05.xx"
                elif s == "磨矿分级":
                    excel_template.worksheets[0]["E" + str(row_num)].value = "2020.06.08~2020.07.25"
                elif s == "钴回收":
                    excel_template.worksheets[0]["E" + str(row_num)].value = "2020.11.29"
                row_num = row_num + 1

                excel_template.worksheets[0]["D" + str(row_num)].value = "电压互感器实验记录"
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                excel_template.worksheets[0]["B" + str(row_num)].value = "--"
                if s == "总降压变电站":
                    excel_template.worksheets[0]["E" + str(row_num)].value = "2020.05.xx"
                elif s == "磨矿分级":
                    excel_template.worksheets[0]["E" + str(row_num)].value = "2020.06.21"
                elif s == "钴回收":
                    excel_template.worksheets[0]["E" + str(row_num)].value = "2020.11.29"
                row_num = row_num + 1
        elif (zifenbu == "电气动力"):
            if s == "磨矿分级":
                excel_template.worksheets[0]["D" + str(row_num)].value = "高压电机实验记录"
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                excel_template.worksheets[0]["B" + str(row_num)].value = "--"
                row_num = row_num + 1
            list_dj = session.query(dian_ji).filter(dian_ji.子项名称 == s).order_by(dian_ji.检查日期).all()
            list_bpq = session.query(bian_pin_qi).filter(bian_pin_qi.子项名称 == s).order_by(bian_pin_qi.检查日期).all()
            list_rq = session.query(ruan_qi).filter(ruan_qi.子项名称 == s).order_by(ruan_qi.检查日期).all()
            if len(list_dj) > 0:
                excel_template.worksheets[0]["D" + str(row_num)].value = "低压电机实验记录"
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                if len(list_dj) == 1 or list_dj[0].检查日期 == list_dj[len(list_dj) - 1].检查日期:
                    excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                        list_dj[0].检查日期, date_formate)
                else:
                    excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                        list_dj[0].检查日期, date_formate) + "~" + datetime.datetime.strftime(
                        list_dj[len(list_dj) - 1].检查日期, date_formate)
                excel_template.worksheets[0]["F" + str(row_num)].value = len(list_dj)
                excel_template.worksheets[0]["B" + str(row_num)].value = "--"
                row_num = row_num + 1
            if len(list_bpq) > 0:
                excel_template.worksheets[0]["D" + str(row_num)].value = "变频器调试记录"
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                if len(list_bpq) == 1 or list_bpq[0].检查日期 == list_bpq[len(list_bpq) - 1].检查日期:
                    excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                        datetime.datetime.strptime(list_bpq[0].检查日期, '%Y-%m-%d'), date_formate)
                else:
                    excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                        datetime.datetime.strptime(list_bpq[0].检查日期, '%Y-%m-%d'),
                        date_formate) + "~" + datetime.datetime.strftime(
                        datetime.datetime.strptime(list_bpq[len(list_bpq) - 1].检查日期, '%Y-%m-%d'), date_formate)
                excel_template.worksheets[0]["F" + str(row_num)].value = len(list_bpq) * 2
                excel_template.worksheets[0]["B" + str(row_num)].value = "--"
                row_num = row_num + 1
            if len(list_rq) > 0:
                excel_template.worksheets[0]["D" + str(row_num)].value = "低压软启实验记录"
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                if len(list_rq) == 1 or list_rq[0].检查日期 == list_rq[len(list_rq) - 1].检查日期:
                    excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                        datetime.datetime.strptime(list_rq[0].检查日期, '%Y-%m-%d'), date_formate)
                else:
                    excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                        datetime.datetime.strptime(list_rq[0].检查日期, '%Y-%m-%d'),
                        date_formate) + "~" + datetime.datetime.strftime(
                        datetime.datetime.strptime(list_rq[len(list_rq) - 1].检查日期, '%Y-%m-%d'), date_formate)
                excel_template.worksheets[0]["F" + str(row_num)].value = len(list_rq)
                excel_template.worksheets[0]["B" + str(row_num)].value = "--"
                row_num = row_num + 1
            excel_template.worksheets[0]["D" + str(row_num)].value = "低压电缆实验记录"
            excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
            excel_template.worksheets[0]["B" + str(row_num)].value = "--"
            row_num = row_num + 1
        elif (zifenbu == "电气照明"):
            list_yinbi_zm = session.query(yin_bi).filter(and_(yin_bi.子项名称 == s, yin_bi.隐蔽类型 == zifenbu)).order_by(
                yin_bi.报验时间).all()
            if len(list_yinbi_zm) == 1:
                excel_template.worksheets[0]["D" + str(row_num)].value = "隐蔽工程验收记录"
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                excel_template.worksheets[0]["B" + str(row_num)].value = fenbu_num + "-05-C5-001"
                excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                    datetime.datetime.strptime(list_yinbi_zm[0].报验时间, '%Y年%m月%d日'), date_formate)
                row_num = row_num + 1
            if len(list_yinbi_zm) > 1:
                excel_template.worksheets[0]["D" + str(row_num)].value = "隐蔽工程验收记录"
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                excel_template.worksheets[0][
                    "B" + str(row_num)].value = fenbu_num + "-05-C5-001~" + "{:0>3d}".format(len(list_yinbi_zm))
                if list_yinbi_zm[0].报验时间 == list_yinbi_zm[len(list_yinbi_zm) - 1].报验时间:
                    excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                        datetime.datetime.strptime(list_yinbi_zm[0].报验时间, '%Y年%m月%d日'), date_formate)
                else:
                    excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                        datetime.datetime.strptime(list_yinbi_zm[0].报验时间, '%Y年%m月%d日'),
                        date_formate) + "~" + datetime.datetime.strftime(
                        datetime.datetime.strptime(list_yinbi_zm[len(list_yinbi_zm) - 1].报验时间, '%Y年%m月%d日'),
                        date_formate)
                excel_template.worksheets[0]["F" + str(row_num)].value = len(list_yinbi_zm)
                row_num = row_num + 1
            list_yinbi_zmsyx = session.query(zhao_ming).filter(and_(zhao_ming.子项名称 == s)).order_by(zhao_ming.开始时间).all()
            if len(list_yinbi_zmsyx) == 1:
                excel_template.worksheets[0]["D" + str(row_num)].value = "建筑照明通电试运行记录"
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                excel_template.worksheets[0]["B" + str(row_num)].value = fenbu_num + "-05-C6-001"
                excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                    datetime.datetime.strptime(list_yinbi_zmsyx[0].开始时间, '%Y-%m-%d'), date_formate)
                row_num = row_num + 1
            if len(list_yinbi_zmsyx) > 1:
                excel_template.worksheets[0]["D" + str(row_num)].value = "建筑照明通电试运行记录"
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                excel_template.worksheets[0]["B" + str(row_num)].value = fenbu_num + "-05-C6-001~" + "{:0>3d}".format(
                    len(list_yinbi_zmsyx))
                if list_yinbi_zmsyx[0].开始时间 == list_yinbi_zmsyx[len(list_yinbi_zmsyx) - 1].开始时间:
                    excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                        datetime.datetime.strptime(list_yinbi_zmsyx[0].开始时间, '%Y-%m-%d'),
                        date_formate)
                else:
                    excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                        datetime.datetime.strptime(list_yinbi_zmsyx[0].开始时间, '%Y-%m-%d'),
                        date_formate) + "~" + datetime.datetime.strftime(
                        datetime.datetime.strptime(list_yinbi_zmsyx[len(list_yinbi_zmsyx) - 1].开始时间, '%Y-%m-%d'),
                        date_formate)
                excel_template.worksheets[0]["F" + str(row_num)].value = len(list_yinbi_zmsyx)
                row_num = row_num + 1
        elif (zifenbu == "防雷及接地"):
            list_yinbi_jd = session.query(yin_bi).filter(and_(yin_bi.子项名称 == s, yin_bi.隐蔽类型 == "防雷接地")).order_by(
                yin_bi.报验时间).all()
            if len(list_yinbi_jd) == 1:
                excel_template.worksheets[0]["D" + str(row_num)].value = "隐蔽工程验收记录"
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                excel_template.worksheets[0]["B" + str(row_num)].value = fenbu_num + "-07-C5-001"
                excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                    datetime.datetime.strptime(list_yinbi_jd[0].报验时间, '%Y年%m月%d日'), date_formate)
                row_num = row_num + 1
            if len(list_yinbi_jd) > 1:
                excel_template.worksheets[0]["D" + str(row_num)].value = "隐蔽工程验收记录"
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                excel_template.worksheets[0]["F" + str(row_num)].value = len(list_yinbi_jd)
                excel_template.worksheets[0][
                    "B" + str(row_num)].value = fenbu_num + "-07-C5-001~" + "{:0>3d}".format(len(list_yinbi_jd))
                if list_yinbi_jd[0].报验时间 == list_yinbi_jd[len(list_yinbi_jd) - 1].报验时间:
                    excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                        datetime.datetime.strptime(list_yinbi_jd[0].报验时间, '%Y年%m月%d日'), date_formate)
                else:
                    excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                        datetime.datetime.strptime(list_yinbi_jd[0].报验时间, '%Y年%m月%d日'),
                        date_formate) + "~" + datetime.datetime.strftime(
                        datetime.datetime.strptime(list_yinbi_jd[len(list_yinbi_jd) - 1].报验时间, '%Y年%m月%d日'),
                        date_formate)

                row_num = row_num + 1
            list_jd = session.query(jie_di_dian_zu).filter(jie_di_dian_zu.子项名称 == s).order_by(jie_di_dian_zu.编号).all()
            if len(list_jd) == 1:
                excel_template.worksheets[0]["D" + str(row_num)].value = "接地电阻测试记录"
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                excel_template.worksheets[0]["B" + str(row_num)].value = list_jd[0].编号
                excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(list_jd[0].实验日期,
                                                                                                    date_formate)
                row_num = row_num + 1
            if len(list_jd) > 1:
                excel_template.worksheets[0]["D" + str(row_num)].value = "接地电阻测试记录"
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                excel_template.worksheets[0]["F" + str(row_num)].value = len(list_jd)
                excel_template.worksheets[0]["B" + str(row_num)].value = list_jd[0].编号 + "~" + "{:0>3d}".format(
                    len(list_jd))
                if list_jd[0].实验日期 == list_jd[len(list_jd) - 1].实验日期:
                    excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(list_jd[0].实验日期,
                                                                                                        date_formate)
                else:
                    excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(list_jd[0].实验日期,
                                                                                                        date_formate) + "~" + datetime.datetime.strftime(
                        list_jd[len(list_jd) - 1].实验日期, date_formate)
                row_num = row_num + 1

        elif (zifenbu == "火灾报警系统"):
            list_yinbi_xf = session.query(yin_bi).filter(and_(yin_bi.子项名称 == s, yin_bi.隐蔽类型 == "火灾报警")).order_by(
                yin_bi.报验时间).all()
            if len(list_yinbi_xf) == 1:
                excel_template.worksheets[0]["D" + str(row_num)].value = "隐蔽工程验收记录"
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                excel_template.worksheets[0]["B" + str(row_num)].value = fenbu_num + "-07-C5-001"
                excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                    datetime.datetime.strptime(list_yinbi_xf[0].报验时间, '%Y年%m月%d日'), date_formate)
                row_num = row_num + 1
            if len(list_yinbi_xf) > 1:
                excel_template.worksheets[0]["D" + str(row_num)].value = "隐蔽工程验收记录"
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                excel_template.worksheets[0]["F" + str(row_num)].value = len(list_yinbi_xf)
                excel_template.worksheets[0][
                    "B" + str(row_num)].value = fenbu_num + "-07-C5-001~" + "{:0>3d}".format(len(list_yinbi_xf))
                excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                    datetime.datetime.strptime(list_yinbi_xf[0].报验时间, '%Y年%m月%d日'), date_formate)
                row_num = row_num + 1
        excel_template.save("E:\工作\庞比\竣工资料\卷内目录\\" + table_name + "\\" + s + ".xlsx")


for i in tqdm(range(len(zixiangs))):
    s = zixiangs[i]
    logger.info(s)
    # 写入起始行
    row_num = 4
    if session.query(jian_yan_pi).filter(jian_yan_pi.子项名称 == s).count() < 1: continue
    export_menu(row_num, zifenbus_xf)
