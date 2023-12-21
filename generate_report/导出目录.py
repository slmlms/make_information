import datetime

from loguru import logger
from sqlalchemy import Column, String, create_engine, Integer, Date
from sqlalchemy import and_
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

import utils.data_util as data

engine = create_engine('mysql+pymysql://root:slmlms123@localhost:3306/jian_yan_pi', echo=False)
Base = declarative_base()
DBSession = sessionmaker(bind=engine)
table_name = "消防"
fenbu_num = "S8"
fenbu_name = "通讯分部工程质量验收记录"
session = DBSession()
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


@logger.catch()
def export_menu(row_num, zifenbus):
    # Excel模板，注意选择打开方式
    excel_template = data.switch_open_excel_template("openpyxl", "inspection_lot/卷内目录 .xlsx")
    # 子项名称
    excel_template.worksheets[0]["A2"].value = "子项名称：" + s
    # 分部
    excel_template.worksheets[0]["B" + str(row_num)].value = "--"
    excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
    excel_template.worksheets[0]["D" + str(row_num)].value = fenbu_name
    row_num = row_num + 1
    for zifenbu in zifenbus:
        list_mb = session.query((mu_ban)).filter(and_(mu_ban.子分部工程名称 == zifenbu, mu_ban.分部工程代号 == fenbu_num)).all()
        assert list_mb is not None
        count = session.query(fenXiang).filter(
            and_(fenXiang.子项名称 == s,
                 fenXiang.检验批数量 > 0, fenXiang.分项工程代号 == list_mb[0].子分部代号)).count()
        # 子分部
        if count > 0:
            excel_template.worksheets[0]["B" + str(row_num)].value = "--"
            excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
            excel_template.worksheets[0]["D" + str(row_num)].value = zifenbu + "子分部工程质量验收记录"
            row_num = row_num + 1
        for lm in list_mb:
            list = session.query(fenXiang).filter(
                and_(fenXiang.分项工程名称 == lm.分项工程名称, fenXiang.子项名称 == s,
                     fenXiang.检验批数量 > 0, fenXiang.分项工程代号 == lm.子分部代号)).order_by(
                fenXiang.分项工程代号).all()
            if len(list) == 0:
                continue
            for l in list:

                logger.info(lm.分项工程名称 + "\t" + lm.子分部工程名称 + "-->" + l.分项工程名称 + "\t" + l.分项工程代号)
                # 分项工程
                excel_template.worksheets[0]["B" + str(row_num)].value = "--"
                excel_template.worksheets[0]["C" + str(row_num)].value = "北方国际"
                excel_template.worksheets[0]["D" + str(row_num)].value = l.分项工程名称 + "分项工程质量验收记录"
                excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                    datetime.datetime.strptime(l.最后一个检验批时间, '%Y-%m-%d'), "%Y.%m.%d")
                row_num = row_num + 1
                logger.info(l.子项名称 + "\t" + l.分项工程名称 + "分项工程质量验收记录")
                for l1 in range(1, int(l.检验批数量) + 1):
                    # 编号
                    bianhao = fenbu_num + "-" + l.分项工程代号.zfill(2) + "-C7-" + "{:0>3d}".format(l1)
                    logger.info(bianhao + "\t" + l.分项工程代号 + "\t" + l.分项工程名称)
                    jyp = session.query(jian_yan_pi).filter(and_(jian_yan_pi.子项名称 == s, jian_yan_pi.分项工程代号 == l.分项工程代号,
                                                                 jian_yan_pi.真编号 == bianhao,
                                                                 jian_yan_pi.分项工程名称 == l.分项工程名称)).first()
                    assert jyp is not None
                    # 检验批写入
                    excel_template.worksheets[0]["B" + str(row_num)].value = bianhao
                    excel_template.worksheets[0]["D" + str(row_num)].value = jyp.分项工程名称 + "检验批质量验收记录"
                    excel_template.worksheets[0]["E" + str(row_num)].value = datetime.datetime.strftime(
                        datetime.datetime.strptime(jyp.真报验时间, '%Y-%m-%d'), "%Y.%m.%d")
                    row_num = row_num + 1
                    logger.info(bianhao + "\t" + jyp.分项工程名称 + "检验批质量验收记录")

                    # print(jyp.分项工程名称,jyp.真报验时间)

                    # print(l.子项名称, l.分项工程名称, bianhao, l.最后一个检验批时间)
        if (zifenbu == "变配电室"):
            excel_template.worksheets[0]["D" + str(row_num)].value = "变压器实验记录"
            excel_template.worksheets[0]["D" + str(row_num + 1)].value = "母线实验记录"
            row_num = row_num + 2
        elif (zifenbu == "电气动力"):
            excel_template.worksheets[0]["D" + str(row_num)].value = "低压电机实验记录"
            excel_template.worksheets[0]["D" + str(row_num + 1)].value = "变频器调试记录"
            excel_template.worksheets[0]["D" + str(row_num + 2)].value = "低压电缆实验记录"
            row_num = row_num + 3
        elif (zifenbu == "电气照明"):
            excel_template.worksheets[0]["D" + str(row_num)].value = "隐蔽工程验收记录"
            excel_template.worksheets[0]["D" + str(row_num + 1)].value = "建筑照明通电试运行记录"
            row_num = row_num + 2
        elif (zifenbu == "防雷及接地"):
            excel_template.worksheets[0]["D" + str(row_num)].value = "隐蔽工程验收记录"
            excel_template.worksheets[0]["D" + str(row_num + 1)].value = "接地电阻测试记录"
            row_num = row_num + 2
        excel_template.save("E:\工作\庞比\竣工资料\卷内目录\\" + table_name + "\\" + s + ".xlsx")


for s in zixiangs:
    # 写入起始行
    row_num = 4
    export_menu(row_num, zifenbus_xf)
