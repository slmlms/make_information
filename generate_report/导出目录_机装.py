from loguru import logger
from rich.console import Console
from sqlalchemy import Column, String, create_engine, Integer, Date
from sqlalchemy import and_, distinct, func
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
from tqdm import tqdm

import utils.data_util as data

# 日期格式
date_formate = "%Y.%m.%d"
zifenbus_jz = ["矿石破碎系统", "胶带输送系统", "给料系统", "磨矿系统", "浸出浓密系统", "电积设备", "萃取设备", "泵送设备", "起重设备", "储槽、储罐", "机械维修设备", "汽车衡",
               "污水处理设备", "非标塔式、圆筒设备", "硫酸标准静设备", "硫酸标准动设备", "发电设备", "絮凝剂制备系统", "闪蒸干燥系统"]
zifenbus_nt = ["除尘系统", "空调系统", "热工系统", "冷却循环系统", "尾气脱硫系统"]
zixiangs = ["粗矿堆A", "粗碎站A", "粗矿堆B", "粗碎站B", "浸出浓密及CCD洗涤", "磨矿分级", "皮带廊及转运站", "铜萃取", "铜电积", "原矿脱水及搅拌浸出", "总降压变电站",
            "给水加压泵房及回水池", "综合管网",
            "新水泵房及输送管线", "硫酸库", "溶剂油库", "选厂设备循环水站", "空压机站", "溶液精滤", "钴回收", "石灰乳及石灰石浆制备", "氧化镁制备"]
engine = create_engine('mysql+pymysql://root:slmlms123@localhost:3306/jian_yan_pi', echo=False)
Base = declarative_base()
DBSession = sessionmaker(bind=engine)
table_name = "暖通"
fenbu_num = "S7"
fenbu_name = table_name + "分部工程质量验收记录"
zifenbus = zifenbus_nt
session = DBSession()
console = Console()


class menu(object):
    wen_jian_bian_hao = None
    ze_ren_zhe = "北方国际"
    wen_jian_ti_ming = None
    ri_qi = None

    def __init__(self, wen_jian_bian_hao, wen_jian_ti_ming, ri_qi):
        self.wen_jian_bian_hao = wen_jian_bian_hao
        self.wen_jian_ti_ming = wen_jian_ti_ming
        self.ri_qi = ri_qi


class ji_zhuang(Base):
    __tablename__ = '机械设备安装'
    id = Column(Integer, primary_key=True)
    子项名称 = Column(String(20))
    分部工程名称 = Column(String(20))
    子分部工程名称 = Column(String(20))
    子分部代号 = Column(String(20))
    分项工程名称 = Column(String(20))
    分项日期 = Column(Date)
    编号 = Column(String(20))
    安装检查日期 = Column(Date)
    检查部位 = Column(String(20))
    电机单试记录 = Column(String(20))
    电机试车日期 = Column(Date)
    试车记录 = Column(String(20))
    设备试车日期 = Column(Date)
    真空箱实验 = Column(String(20))
    真空箱日期 = Column(Date)
    煤油渗漏实验 = Column(String(20))
    煤油日期 = Column(Date)
    满水实验 = Column(String(20))
    满水日期 = Column(Date)


class fen_bu_zi_fen_bu(Base):
    __tablename__ = '分部子分部'
    id = Column(Integer, primary_key=True)
    子项名称 = Column(String(20))
    分部名称 = Column(String(20))
    日期 = Column(Date)


for i in tqdm(range(len(zixiangs))):
    zixiang = zixiangs[i]
    logger.info(zixiang)
    # zixiang.set_discription("当前进度: %s", i)
    list_row = []
    # Excel模板，注意选择打开方式
    excel_template = data.switch_open_excel_template("openpyxl", "inspection_lot/卷内目录 .xlsx")
    # 子项名称
    excel_template.worksheets[0]["A2"].value = "子项名称：" + zixiang
    # 写入起始行
    row_num = 4
    if session.query(ji_zhuang.id).filter(
            and_(ji_zhuang.子项名称 == zixiang, ji_zhuang.分部工程名称 == table_name)).count() < 1: continue
    # 分部工程
    fenbu_date = session.query(func.date_format(fen_bu_zi_fen_bu.日期, date_formate)).filter(
        and_(fen_bu_zi_fen_bu.子项名称 == zixiang, fen_bu_zi_fen_bu.分部名称 == fenbu_name)).first()
    if fenbu_date is None: fenbu_date = [""]
    list_row.append(menu(wen_jian_bian_hao="--", wen_jian_ti_ming=fenbu_name, ri_qi=fenbu_date[0]))

    for zifenbu in zifenbus:

        list_fx = session.query(distinct(ji_zhuang.分项工程名称)).filter(
            and_(ji_zhuang.子项名称 == zixiang, ji_zhuang.子分部工程名称 == zifenbu)).all()
        if len(list_fx) < 1: continue
        # 子分部工程
        zifenbu_date = session.query(func.date_format(fen_bu_zi_fen_bu.日期, date_formate)).filter(
            and_(fen_bu_zi_fen_bu.子项名称 == zixiang, fen_bu_zi_fen_bu.分部名称 == zifenbu + "子分部工程质量验收记录")).first()
        if zifenbu_date is None: zifenbu_date = [""]
        list_row.append(menu("--", zifenbu + "子分部工程质量验收记录", zifenbu_date[0]))
        if len(list_fx) > 0:
            for fx in list_fx:
                # 分项
                tiao_jian = and_(ji_zhuang.子项名称 == zixiang, ji_zhuang.子分部工程名称 == zifenbu, ji_zhuang.分项工程名称 == fx[0])
                fx_date = session.query(func.date_format(ji_zhuang.分项日期, date_formate)).filter(tiao_jian).first()[0]
                list_row.append(menu("--", fx[0] + "分项工程质量验收记录", fx_date))
                list_jz = session.query(ji_zhuang).filter(tiao_jian).all()
                for jx in list_jz:
                    tiao_jian_jyp = and_(ji_zhuang.子项名称 == zixiang, ji_zhuang.子分部工程名称 == zifenbu,
                                         ji_zhuang.分项工程名称 == fx[0], ji_zhuang.检查部位 == jx.检查部位)
                    if jx.编号 is not None:
                        list_row.append(
                            menu(jx.编号, "设备安装检查记录(" + jx.检查部位 + ")",
                                 session.query(func.date_format(ji_zhuang.安装检查日期, date_formate)).filter(
                                     tiao_jian_jyp).first()[0]))
                    if jx.电机单试记录 is not None:
                        list_row.append(
                            menu(jx.电机单试记录, "电动机单体试车记录(" + jx.检查部位 + ")",
                                 session.query(func.date_format(ji_zhuang.电机试车日期, date_formate)).filter(
                                     tiao_jian_jyp).first()[0]))
                    if jx.试车记录 is not None:
                        list_row.append(
                            menu(jx.试车记录, "机械设备单体试车记录(" + jx.检查部位 + ")",
                                 session.query(func.date_format(ji_zhuang.设备试车日期, date_formate)).filter(
                                     tiao_jian_jyp).first()[0]))
                    if jx.真空箱实验 is not None:
                        list_row.append(
                            menu(jx.真空箱实验, "真空箱实验(" + jx.检查部位 + ")",
                                 session.query(func.date_format(ji_zhuang.真空箱日期, date_formate)).filter(
                                     tiao_jian_jyp).first()[0]))
                    if jx.煤油渗漏实验 is not None:
                        list_row.append(
                            menu(jx.煤油渗漏实验, "煤油渗漏实验(" + jx.检查部位 + ")",
                                 session.query(func.date_format(ji_zhuang.煤油日期, date_formate)).filter(
                                     tiao_jian_jyp).first()[0]))
                    if jx.满水实验 is not None:
                        list_row.append(
                            menu(jx.满水实验, "满水实验(" + jx.检查部位 + ")",
                                 session.query(func.date_format(ji_zhuang.满水日期, date_formate)).filter(
                                     tiao_jian_jyp).first()[0]))
    for l in list_row:
        # print(l.wen_jian_bian_hao + "--" + l.wen_jian_ti_ming + "--" + l.ri_qi)
        excel_template.worksheets[0]["B" + str(row_num)].value = l.wen_jian_bian_hao
        excel_template.worksheets[0]["C" + str(row_num)].value = l.ze_ren_zhe
        excel_template.worksheets[0]["D" + str(row_num)].value = l.wen_jian_ti_ming
        excel_template.worksheets[0]["E" + str(row_num)].value = l.ri_qi
        row_num = row_num + 1
    excel_template.save("E:\工作\庞比\竣工资料\卷内目录\\" + table_name + "\\" + zixiang + ".xlsx")
