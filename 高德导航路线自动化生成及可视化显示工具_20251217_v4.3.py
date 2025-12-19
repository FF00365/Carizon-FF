import sys
import os
import json
import pandas as pd
import requests
import time
import folium
from folium.plugins import MarkerCluster, AntPath
import math
from urllib.parse import urlparse, parse_qs, quote
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QFileDialog, 
                            QLabel, QVBoxLayout, QHBoxLayout, QGridLayout, QWidget, QProgressBar, 
                            QMessageBox, QTextEdit, QLineEdit, QTabWidget, QGroupBox,
                            QButtonGroup, QRadioButton, QListWidget, QListWidgetItem, QComboBox, QDialog, QCheckBox, QSplitter,
                            QSpinBox, QDoubleSpinBox, QTreeWidget, QTreeWidgetItem, QSizePolicy)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon
from shapely.geometry import LineString
import geopandas as gpd
import webbrowser
import logging
import threading
import tempfile
import random
import re
import hashlib

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 版本信息
VERSION = "V4.0"

# 颜色列表，用于不同路线，去掉不醒目的颜色
COLORS = ['red',  'green', 'purple', 'darkred', 
           'crimson', 'magenta', 'brown','darkblue',
          'maroon', 'forestgreen',  'hotpink',  'fuchsia', 'coral',  'violet']
          #'darkgreen','indigo', 'dodgerblue'

# 全国主要城市及行政区数据（新增）
CITY_DISTRICTS = {
    # 直辖市
    "北京": ["东城区", "西城区", "朝阳区", "丰台区", "石景山区", "海淀区", "门头沟区", "房山区", "通州区", "顺义区", "昌平区", "大兴区", "怀柔区", "平谷区", "密云区", "延庆区"],
    "上海": ["黄浦区", "徐汇区", "长宁区", "静安区", "普陀区", "虹口区", "杨浦区", "闵行区", "宝山区", "嘉定区", "浦东新区", "金山区", "松江区", "青浦区", "奉贤区", "崇明区"],
    "天津": ["和平区", "河东区", "河西区", "南开区", "河北区", "红桥区", "东丽区", "西青区", "津南区", "北辰区", "武清区", "宝坻区", "滨海新区", "宁河区", "静海区", "蓟州区"],
    "重庆": ["万州区", "涪陵区", "渝中区", "大渡口区", "江北区", "沙坪坝区", "九龙坡区", "南岸区", "北碚区", "綦江区", "大足区", "渝北区", "巴南区", "黔江区", "长寿区", "江津区", "合川区", "永川区", "南川区", "璧山区", "铜梁区", "潼南区", "荣昌区", "开州区", "梁平区", "武隆区", "城口县", "丰都县", "垫江县", "忠县", "云阳县", "奉节县", "巫山县", "巫溪县", "石柱土家族自治县", "秀山土家族苗族自治县", "酉阳土家族苗族自治县", "彭水苗族土家族自治县"],
    
    # 省会城市
    "广州": ["越秀区", "海珠区", "荔湾区", "天河区", "白云区", "黄埔区", "番禺区", "花都区", "南沙区", "从化区", "增城区"],
    "深圳": ["罗湖区", "福田区", "南山区", "宝安区", "龙岗区", "盐田区", "龙华区", "坪山区", "光明区", "大鹏新区"],
    "成都": ["锦江区", "青羊区", "金牛区", "武侯区", "成华区", "龙泉驿区", "青白江区", "新都区", "温江区", "双流区", "郫都区", "新津区", "都江堰市", "彭州市", "邛崃市", "崇州市", "简阳市", "金堂县", "大邑县", "蒲江县"],
    "杭州": ["上城区", "下城区", "江干区", "拱墅区", "西湖区", "滨江区", "萧山区", "余杭区", "富阳区", "临安区", "建德市", "桐庐县", "淳安县"],
    "武汉": ["江岸区", "江汉区", "硚口区", "汉阳区", "武昌区", "青山区", "洪山区", "东西湖区", "汉南区", "蔡甸区", "江夏区", "黄陂区", "新洲区"],
    "西安": ["新城区", "碑林区", "莲湖区", "灞桥区", "未央区", "雁塔区", "阎良区", "临潼区", "长安区", "高陵区", "鄠邑区", "蓝田县", "周至县"],
    "南京": ["玄武区", "秦淮区", "建邺区", "鼓楼区", "浦口区", "栖霞区", "雨花台区", "江宁区", "六合区", "溧水区", "高淳区"],
    "长沙": ["芙蓉区", "天心区", "岳麓区", "开福区", "雨花区", "望城区", "宁乡市", "浏阳市", "长沙县"],
    "郑州": ["中原区", "二七区", "管城回族区", "金水区", "上街区", "惠济区", "巩义市", "荥阳市", "新密市", "新郑市", "登封市", "中牟县"],
    "济南": ["历下区", "市中区", "槐荫区", "天桥区", "历城区", "长清区", "章丘区", "济阳区", "莱芜区", "钢城区", "平阴县", "商河县"],
    "青岛": ["市南区", "市北区", "黄岛区", "崂山区", "李沧区", "城阳区", "即墨区", "胶州市", "平度市", "莱西市"],
    "沈阳": ["和平区", "沈河区", "大东区", "皇姑区", "铁西区", "苏家屯区", "浑南区", "沈北新区", "于洪区", "辽中区", "新民市", "康平县", "法库县"],
    "大连": ["中山区", "西岗区", "沙河口区", "甘井子区", "旅顺口区", "金州区", "普兰店区", "瓦房店市", "庄河市", "长海县"],
    "哈尔滨": ["道里区", "南岗区", "道外区", "平房区", "松北区", "香坊区", "呼兰区", "阿城区", "双城区", "尚志市", "五常市", "依兰县", "方正县", "宾县", "巴彦县", "木兰县", "通河县", "延寿县"],
    "长春": ["南关区", "宽城区", "朝阳区", "二道区", "绿园区", "双阳区", "九台区", "榆树市", "德惠市", "公主岭市", "农安县"],
    "福州": ["鼓楼区", "台江区", "仓山区", "晋安区", "马尾区", "长乐区", "福清市", "闽侯县", "连江县", "罗源县", "闽清县", "永泰县", "平潭县"],
    "厦门": ["思明区", "海沧区", "湖里区", "集美区", "同安区", "翔安区"],
    "昆明": ["五华区", "盘龙区", "官渡区", "西山区", "东川区", "呈贡区", "晋宁区", "安宁市", "富民县", "宜良县", "石林彝族自治县", "嵩明县", "禄劝彝族苗族自治县", "寻甸回族彝族自治县"],
    "南昌": ["东湖区", "西湖区", "青云谱区", "湾里区", "青山湖区", "新建区", "南昌县", "安义县", "进贤县"],
    "贵阳": ["南明区", "云岩区", "花溪区", "乌当区", "白云区", "观山湖区", "清镇市", "开阳县", "息烽县", "修文县"],
    "南宁": ["兴宁区", "青秀区", "江南区", "西乡塘区", "良庆区", "邕宁区", "武鸣区", "隆安县", "马山县", "上林县", "宾阳县", "横州市"],
    "拉萨": ["城关区", "堆龙德庆区", "达孜区", "林周县", "当雄县", "尼木县", "曲水县", "墨竹工卡县"],
    "西宁": ["城东区", "城中区", "城西区", "城北区", "大通回族土族自治县", "湟中县", "湟源县"],
    "兰州": ["城关区", "七里河区", "西固区", "安宁区", "红古区", "永登县", "皋兰县", "榆中县"],
    "银川": ["兴庆区", "西夏区", "金凤区", "灵武市", "永宁县", "贺兰县"],
    "乌鲁木齐": ["天山区", "沙依巴克区", "新市区", "水磨沟区", "头屯河区", "达坂城区", "米东区", "乌鲁木齐县"],
    "呼和浩特": ["新城区", "回民区", "玉泉区", "赛罕区", "土默特左旗", "托克托县", "和林格尔县", "清水河县", "武川县"],
    "太原": ["小店区", "迎泽区", "杏花岭区", "尖草坪区", "万柏林区", "晋源区", "清徐县", "阳曲县", "娄烦县", "古交市"],
    "石家庄": ["长安区", "桥西区", "新华区", "井陉矿区", "裕华区", "藁城区", "鹿泉区", "栾城区", "辛集市", "晋州市", "新乐市", "井陉县", "正定县", "行唐县", "灵寿县", "高邑县", "深泽县", "赞皇县", "无极县", "平山县", "元氏县", "赵县"],
    "合肥": ["瑶海区", "庐阳区", "蜀山区", "包河区", "长丰县", "肥东县", "肥西县", "庐江县", "巢湖市"],
    "南京": ["玄武区", "秦淮区", "建邺区", "鼓楼区", "浦口区", "栖霞区", "雨花台区", "江宁区", "六合区", "溧水区", "高淳区"],
    "杭州": ["上城区", "下城区", "江干区", "拱墅区", "西湖区", "滨江区", "萧山区", "余杭区", "富阳区", "临安区", "建德市", "桐庐县", "淳安县"],
    "宁波": ["海曙区", "江北区", "北仑区", "镇海区", "鄞州区", "奉化区", "余姚市", "慈溪市", "象山县", "宁海县"],
    "苏州": ["姑苏区", "虎丘区", "吴中区", "相城区", "吴江区", "苏州工业园区", "常熟市", "张家港市", "昆山市", "太仓市"],
    "无锡": ["锡山区", "惠山区", "滨湖区", "梁溪区", "新吴区", "江阴市", "宜兴市"],
    "常州": ["天宁区", "钟楼区", "新北区", "武进区", "金坛区", "溧阳市"],
    "徐州": ["云龙区", "鼓楼区", "贾汪区", "泉山区", "铜山区", "丰县", "沛县", "睢宁县", "新沂市", "邳州市"],
    "南通": ["崇川区", "港闸区", "通州区", "如东县", "启东市", "如皋市", "海门市", "海安市"],
    "扬州": ["广陵区", "邗江区", "江都区", "宝应县", "仪征市", "高邮市"],
    "镇江": ["京口区", "润州区", "丹徒区", "丹阳市", "扬中市", "句容市"],
    "泰州": ["海陵区", "高港区", "姜堰区", "兴化市", "靖江市", "泰兴市"],
    "盐城": ["亭湖区", "盐都区", "大丰区", "响水县", "滨海县", "阜宁县", "射阳县", "建湖县", "东台市"],
    "淮安": ["淮安区", "淮阴区", "清江浦区", "洪泽区", "涟水县", "盱眙县", "金湖县"],
    "连云港": ["连云区", "海州区", "赣榆区", "东海县", "灌云县", "灌南县"],
    "宿迁": ["宿城区", "宿豫区", "沭阳县", "泗阳县", "泗洪县"],
    
    # 其他主要城市
    "东莞": ["莞城街道", "南城街道", "东城街道", "万江街道", "石碣镇", "石龙镇", "茶山镇", "石排镇", "企石镇", "横沥镇", "桥头镇", "谢岗镇", "东坑镇", "常平镇", "寮步镇", "樟木头镇", "大朗镇", "黄江镇", "清溪镇", "塘厦镇", "凤岗镇", "大岭山镇", "长安镇", "虎门镇", "厚街镇", "沙田镇", "道滘镇", "洪梅镇", "麻涌镇", "望牛墩镇", "中堂镇", "高埗镇"],
    "佛山": ["禅城区", "南海区", "顺德区", "三水区", "高明区"],
    "中山": ["石岐区", "东区", "西区", "南区", "五桂山区", "火炬开发区", "黄圃镇", "南头镇", "东凤镇", "阜沙镇", "小榄镇", "东升镇", "古镇镇", "横栏镇", "三角镇", "民众镇", "南朗镇", "港口镇", "大涌镇", "沙溪镇", "三乡镇", "板芙镇", "神湾镇", "坦洲镇"],
    "珠海": ["香洲区", "斗门区", "金湾区"],
    "惠州": ["惠城区", "惠阳区", "惠东县", "博罗县", "龙门县", "大亚湾经济技术开发区", "仲恺高新技术产业开发区"],
    "汕头": ["龙湖区", "金平区", "濠江区", "潮阳区", "潮南区", "澄海区", "南澳县"],
    "江门": ["蓬江区", "江海区", "新会区", "台山市", "开平市", "鹤山市", "恩平市"],
    "湛江": ["赤坎区", "霞山区", "坡头区", "麻章区", "廉江市", "雷州市", "吴川市", "遂溪县", "徐闻县"],
    "茂名": ["茂南区", "电白区", "高州市", "化州市", "信宜市"],
    "肇庆": ["端州区", "鼎湖区", "高要区", "四会市", "广宁县", "怀集县", "封开县", "德庆县"],
    "揭阳": ["榕城区", "揭东区", "揭西县", "惠来县", "普宁市"],
    "潮州": ["湘桥区", "潮安区", "饶平县"],
    "汕尾": ["城区", "海丰县", "陆河县", "陆丰市"],
    "河源": ["源城区", "紫金县", "龙川县", "连平县", "和平县", "东源县"],
    "阳江": ["江城区", "阳东区", "阳西县", "阳春市"],
    "清远": ["清城区", "清新区", "英德市", "连州市", "佛冈县", "阳山县", "连山壮族瑶族自治县", "连南瑶族自治县"],
    "韶关": ["武江区", "浈江区", "曲江区", "乐昌市", "南雄市", "始兴县", "仁化县", "翁源县", "乳源瑶族自治县", "新丰县"],
    "梅州": ["梅江区", "梅县区", "兴宁市", "大埔县", "丰顺县", "五华县", "平远县", "蕉岭县"],
    "云浮": ["云城区", "云安区", "罗定市", "新兴县", "郁南县"],
    "柳州": ["城中区", "鱼峰区", "柳南区", "柳北区", "柳江区", "柳城县", "鹿寨县", "融安县", "融水苗族自治县", "三江侗族自治县"],
    "桂林": ["秀峰区", "叠彩区", "象山区", "七星区", "雁山区", "临桂区", "阳朔县", "灵川县", "全州县", "兴安县", "永福县", "灌阳县", "龙胜各族自治县", "资源县", "平乐县", "恭城瑶族自治县"],
    "梧州": ["万秀区", "长洲区", "龙圩区", "苍梧县", "藤县", "蒙山县", "岑溪市"],
    "北海": ["海城区", "银海区", "铁山港区", "合浦县"],
    "防城港": ["港口区", "防城区", "上思县", "东兴市"],
    "钦州": ["钦南区", "钦北区", "灵山县", "浦北县"],
    "贵港": ["港北区", "港南区", "覃塘区", "平南县", "桂平市"],
    "玉林": ["玉州区", "福绵区", "容县", "陆川县", "博白县", "兴业县", "北流市"],
    "百色": ["右江区", "田阳区", "田东县", "德保县", "那坡县", "凌云县", "乐业县", "田林县", "西林县", "隆林各族自治县", "靖西市", "平果市"],
    "贺州": ["八步区", "平桂区", "昭平县", "钟山县", "富川瑶族自治县"],
    "河池": ["金城江区", "宜州区", "南丹县", "天峨县", "凤山县", "东兰县", "罗城仫佬族自治县", "环江毛南族自治县", "巴马瑶族自治县", "都安瑶族自治县", "大化瑶族自治县"],
    "来宾": ["兴宾区", "忻城县", "象州县", "武宣县", "金秀瑶族自治县", "合山市"],
    "崇左": ["江州区", "扶绥县", "宁明县", "龙州县", "大新县", "天等县", "凭祥市"]
}

# GCJ-02坐标转换工具函数
# def gcj02_from_wgs84(lng, lat):
#     """将WGS84坐标转换为GCJ-02坐标(高德坐标系)
#     :param lng: WGS84坐标系的经度
#     :param lat: WGS84坐标系的纬度
#     :return: 转换后的GCJ-02坐标(lng, lat)
#     """
#     if out_of_china(lng, lat):
#         return lng, lat
#     dlat = _transformlat(lng - 105.0, lat - 35.0)
#     dlng = _transformlng(lng - 105.0, lat - 35.0)
#     radlat = lat / 180.0 * math.pi
#     magic = math.sin(radlat)
#     magic = 1 - 0.00669342162296594323 * magic * magic
#     sqrtmagic = math.sqrt(magic)
#     dlat = (dlat * 180.0) / ((6335552.3142451795 * (1 - 0.00669342162296594323)) / (magic * sqrtmagic) * math.pi)
#     dlng = (dlng * 180.0) / (6378245.0 / sqrtmagic * math.cos(radlat) * math.pi)
#     gcj_lat = lat + dlat
#     gcj_lng = lng + dlng
#     return gcj_lng, gcj_lat


def out_of_china(lng, lat):
    """判断坐标是否在中国大陆之外"""
    return not (73.66 < lng < 135.05 and 3.86 < lat < 53.55)


def _transformlat(lng, lat):
    ret = -100.0 + 2.0 * lng + 3.0 * lat + 0.2 * lat * lat + 0.1 * lng * lat + 0.2 * math.sqrt(math.fabs(lng))
    ret += (20.0 * math.sin(6.0 * lng * math.pi) + 20.0 * math.sin(2.0 * lng * math.pi)) * 2.0 / 3.0
    ret += (20.0 * math.sin(lat * math.pi) + 40.0 * math.sin(lat / 3.0 * math.pi)) * 2.0 / 3.0
    ret += (160.0 * math.sin(lat / 12.0 * math.pi) + 320 * math.sin(lat * math.pi / 30.0)) * 2.0 / 3.0
    return ret


def _transformlng(lng, lat):
    ret = 300.0 + lng + 2.0 * lat + 0.1 * lng * lng + 0.1 * lng * lat + 0.1 * math.sqrt(math.fabs(lng))
    ret += (20.0 * math.sin(6.0 * lng * math.pi) + 20.0 * math.sin(2.0 * lng * math.pi)) * 2.0 / 3.0
    ret += (20.0 * math.sin(lng * math.pi) + 40.0 * math.sin(lng / 3.0 * math.pi)) * 2.0 / 3.0
    ret += (150.0 * math.sin(lng / 12.0 * math.pi) + 300.0 * math.sin(lng / 30.0 * math.pi)) * 2.0 / 3.0
    return ret

class RouteCalculator(QThread):
    """线程类，用于计算路线，避免UI卡顿"""
    progress_updated = pyqtSignal(int)
    calculation_finished = pyqtSignal(list, list, list, list)
    error_occurred = pyqtSignal(str)

    def __init__(self, waypoints, key, backup_keys=None):
        super().__init__()
        self.waypoints = waypoints
        self.key = key
        self.backup_keys = backup_keys or []
        self.current_key_index = -1  # 从主密钥开始
        
    # def gcj02_to_wgs84(self, gcj_lon, gcj_lat):
    #     """
    #     GCJ-02(火星坐标系)转WGS84
    #     :param gcj_lon: GCJ-02经度
    #     :param gcj_lat: GCJ-02纬度
    #     :return: WGS84坐标(lon, lat)
    #     """
    #     # 判断是否在国内
    #     if not (72.004 <= gcj_lon <= 137.8347 and 0.8293 <= gcj_lat <= 55.8271):
    #         return gcj_lon, gcj_lat
            
    #     a = 6378245.0
    #     ee = 0.00669342162296594323
        
    #     # 转换GCJ-02到WGS84
    #     dlat = self._transform_lat(gcj_lon - 105.0, gcj_lat - 35.0)
    #     dlon = self._transform_lon(gcj_lon - 105.0, gcj_lat - 35.0)
    #     radlat = gcj_lat / 180.0 * math.pi
    #     magic = math.sin(radlat)
    #     magic = 1 - ee * magic * magic
    #     sqrtmagic = math.sqrt(magic)
    #     dlat = (dlat * 180.0) / ((a * (1 - ee)) / (magic * sqrtmagic) * math.pi)
    #     dlon = (dlon * 180.0) / (a / sqrtmagic * math.cos(radlat) * math.pi)
    #     wgs_lat = gcj_lat - dlat
    #     wgs_lon = gcj_lon - dlon
        
    #     return wgs_lon, wgs_lat
        
    def _transform_lat(self, x, y):
        ret = -100.0 + 2.0 * x + 3.0 * y + 0.2 * y * y + 0.1 * x * y + 0.2 * math.sqrt(abs(x))
        ret += (20.0 * math.sin(6.0 * x * math.pi) + 20.0 * math.sin(2.0 * x * math.pi)) * 2.0 / 3.0
        ret += (20.0 * math.sin(y * math.pi) + 40.0 * math.sin(y / 3.0 * math.pi)) * 2.0 / 3.0
        ret += (160.0 * math.sin(y / 12.0 * math.pi) + 320 * math.sin(y * math.pi / 30.0)) * 2.0 / 3.0
        return ret
    
    def _transform_lon(self, x, y):
        ret = 300.0 + x + 2.0 * y + 0.1 * x * x + 0.1 * x * y + 0.1 * math.sqrt(abs(x))
        ret += (20.0 * math.sin(6.0 * x * math.pi) + 20.0 * math.sin(2.0 * x * math.pi)) * 2.0 / 3.0
        ret += (20.0 * math.sin(x * math.pi) + 40.0 * math.sin(x / 3.0 * math.pi)) * 2.0 / 3.0
        ret += (150.0 * math.sin(x / 12.0 * math.pi) + 300.0 * math.sin(x / 30.0 * math.pi)) * 2.0 / 3.0
        return ret

    def get_next_key(self):
        """获取下一个API密钥"""
        self.current_key_index += 1
        if self.current_key_index == 0:
            return self.key  # 返回主密钥
        elif self.current_key_index <= len(self.backup_keys):
            return self.backup_keys[self.current_key_index - 1]  # 返回备用密钥
        else:
            return None  # 所有密钥都尝试过了

    def run(self):
        try:
            all_coordinates = []
            all_road_types = []  # 存储所有路段的道路类型
            all_road_names = []  # 存储所有路段的道路名称
            route_segments = []
            total_points = len(self.waypoints)
            
            for i in range(total_points - 1):
                # 更新进度
                progress = int((i / (total_points - 1)) * 100)
                self.progress_updated.emit(progress)
                
                origin = self.waypoints[i]
                destination = self.waypoints[i+1]
                
                # 获取路线
                try:
                    coordinates, road_types, road_names = self.get_route(origin, destination)
                except Exception as e:
                    # 如果使用当前密钥失败，尝试使用备用密钥
                    next_key = self.get_next_key()
                    if next_key:
                        print(f"\n当前密钥失败，尝试使用备用密钥: {next_key}")
                        self.key = next_key
                        coordinates, road_types, road_names = self.get_route(origin, destination)
                    else:
                        raise e  # 所有密钥都失败了，抛出异常
                
                # 记录路段信息
                segment_info = {
                    "起点": f"点{i+1}",
                    "终点": f"点{i+2}",
                    "起点坐标": origin,
                    "终点坐标": destination,
                    "坐标点数": len(coordinates),
                    "道路类型": road_types,
                    "道路名称": road_names
                }
                route_segments.append(segment_info)
                
                # 添加到总路线
                all_coordinates.extend(coordinates)
                all_road_types.extend(road_types)  # 添加道路类型信息
                all_road_names.extend(road_names)  # 添加道路名称信息
                
                # 添加延迟，避免API配额限制
                time.sleep(1)
            
            self.progress_updated.emit(100)
            self.calculation_finished.emit(all_coordinates, route_segments, all_road_types, all_road_names)  # 添加道路名称信息
            
        except Exception as e:
            self.error_occurred.emit(str(e))

    def get_route(self, origin, destination):
        """获取从起点到终点的驾车路线"""
        # 尝试使用V5版本API
        url_v5 = f"https://restapi.amap.com/v5/direction/driving?origin={origin}&destination={destination}&key={self.key}&show_fields=cost,polyline,road_type"
        
        # 尝试使用V3版本API（作为备用）
        url_v3 = f"https://restapi.amap.com/v3/direction/driving?origin={origin}&destination={destination}&key={self.key}&extensions=all"
        
        print(f"\n=== 请求URL (V5) ===\n{url_v5}")
        
        # 添加重试机制
        max_retries = 3
        retry_delay = 2
        
        for attempt in range(max_retries):
            try:
                # 首先尝试V5 API
                response = requests.get(url_v5)
                data = response.json()
                
                # 检查V5 API是否成功
                if data.get('status') != '1' or 'route' not in data or 'paths' not in data['route'] or not data['route']['paths']:
                    print(f"\n=== V5 API失败，尝试V3 API ===")
                    # 如果V5 API失败，尝试V3 API
                    response = requests.get(url_v3)
                    data = response.json()
                    print(f"\n=== V3 API响应 ===\n{str(response.text)[:1000]}")
                    
                    # 检查V3 API是否成功
                    if data.get('status') != '1' or 'route' not in data or 'paths' not in data['route'] or not data['route']['paths']:
                        if attempt < max_retries - 1:
                            time.sleep(retry_delay * (attempt + 1))
                            continue
                        else:
                            raise Exception(f"无法获取路线，请检查输入参数或API Key: {data}")
                
                # 调试：打印完整的原始API响应（前1000个字符）
                print(f"\n=== 原始API响应（前1000个字符）===\n{str(response.text)[:1000]}")
                
                # 调试：打印完整的API响应
                print(f"\n=== API响应数据 ===")
                print(f"状态码: {data.get('status')}")
                print(f"信息码: {data.get('infocode')}")
                print(f"信息: {data.get('info')}")
                
                if 'route' in data and 'paths' in data['route'] and len(data['route']['paths']) > 0:
                    path = data['route']['paths'][0]
                    print(f"路径距离: {path.get('distance')}米")
                    print(f"路径时间: {path.get('duration')}秒")
                    print(f"路段数量: {len(path.get('steps', []))}")
                    
                    # 打印第一个路段的详细信息
                    if 'steps' in path and len(path['steps']) > 0:
                        first_step = path['steps'][0]
                        print("\n第一个路段详细信息:")
                        for key, value in first_step.items():
                            if key != 'polyline':  # polyline太长，不打印
                                print(f"  {key}: {value}")
                        
                        # 特别检查road_type字段
                        if 'road_type' in first_step:
                            print(f"\n特别注意: road_type存在，值为 '{first_step['road_type']}'")
                        else:
                            print(f"\n特别注意: road_type字段不存在！")
                            
                            # 如果是V3 API，可能使用不同的字段名
                            if 'highway' in first_step:
                                print(f"  发现highway字段: {first_step['highway']}")
                            if 'toll' in first_step:
                                print(f"  发现toll字段: {first_step['toll']}")
                
                if data['status'] == '1':
                    route = data['route']['paths'][0]  # 获取第一条路线的信息
                    steps = route['steps']
                    # 提取经纬度点
                    coordinates = []
                    road_types = []  # 存储道路类型信息
                    road_names = []  # 存储道路名称信息
                    
                    # 转换起点和终点坐标
                    #origin_lon, origin_lat = map(float, origin.split(','))
                    #dest_lon, dest_lat = map(float, destination.split(','))

                    #origin_lon, origin_lat = self.origin_lon, origin_lat
                    #origin_lon, origin_lat = self.gcj02_to_wgs84(origin_lon, origin_lat)
                    #dest_lon, dest_lat = self.dest_lon, dest_lat
                    #dest_lon, dest_lat = self.gcj02_to_wgs84(dest_lon, dest_lat)
                    
                    # 检查steps中是否有road_type字段
                    has_road_type = False
                    has_highway = False
                    for step in steps:
                        if 'road_type' in step:
                            has_road_type = True
                        if 'highway' in step:
                            has_highway = True
                    
                    if not has_road_type:
                        print("\n警告：API返回的steps中没有road_type字段！")
                        if has_highway:
                            print("但发现了highway字段，将使用它来确定高速公路")
                    
                    # 高速公路和高架道路的关键词
                    highway_keywords = [
                        # 通用高速关键词
                        '高速', '高速公路', '高速路', '高速环', '环高速', '机场高速', '枢纽', '互通',
                        # 国道高速编号前缀
                        'G', 'S', '国道高速', '省道高速', '高速国道', '高速省道',
                        # 京津冀及周边高速
                        '京藏高速', '京港澳高速', '京沪高速', '京津高速', '京昆高速', '京开高速', '京承高速', '京台高速',
                        '京哈高速', '京礼高速', '京新高速', '京张高速', '京石高速', '京秦高速', '大广高速', '唐津高速',
                        '津石高速', '津晋高速', '荣乌高速', '青银高速', '石太高速', '石黄高速', '保沧高速',
                        # 长三角高速
                        '沪宁高速', '沪杭高速', '沪蓉高速', '沪渝高速', '沪陕高速', '沪昆高速', '杭甬高速', '宁杭高速',
                        '杭州绕城高速', '南京绕城高速', '苏通高速', '苏嘉杭高速', '宁常高速', '常台高速', '锡宜高速',
                        '沿江高速', '沪常高速', '常嘉高速', '嘉绍高速', '杭金衢高速', '申嘉湖高速', '湖杭高速',
                        # 珠三角高速
                        '广深高速', '广澳高速', '广惠高速', '广河高速', '广州绕城高速', '深圳绕城高速', '莞深高速',
                        '虎门高速', '广珠西高速', '广珠东高速', '佛开高速', '佛山一环高速', '珠三角环线高速',
                        '深汕高速', '惠盐高速', '厦深高速', '汕湛高速',
                        # 东北地区高速
                        '沈大高速', '长深高速', '哈大高速', '哈齐高速', '丹阜高速', '沈丹高速', '沈海高速',
                        '长春绕城高速', '沈阳绕城高速', '哈尔滨绕城高速', '鹤大高速', '大广高速',
                        # 中西部高速
                        '成渝高速', '成雅高速', '成绵高速', '绵西高速', '成灌高速', '成温邛高速', '渝湘高速',
                        '渝黔高速', '兰海高速', '西汉高速', '福银高速', '沪渝高速', '沪陕高速', '连霍高速',
                        '青兰高速', '银川绕城高速', '西安绕城高速', '兰州绕城高速', '成都绕城高速', '重庆绕城高速',
                        '长株潭环线高速', '武汉城市圈环线高速',
                        # 其他主要高速
                        '长深高速', '长吉高速', '长张高速', '济广高速', '济青高速', '济南绕城高速', '青岛绕城高速',
                        '日兰高速', '胶州湾高速', '杭州湾环线高速', '杭州湾跨海大桥',
                        # 高速收费站和服务区
                        '收费站', '服务区', '高速出口', '高速入口', 'IC', 'JCT'
                    ]
                    
                    elevated_keywords = [
                        # 高架道路关键词
                        '高架', '高架路', '高架桥', '立交', '立交桥', '快速路', '快速干道', 
                        '城市快速路', '城市快速', '快速通道', '高架道路', '高架通道', '高架环路',
                        '内环高架', '中环高架', '外环高架', '高架环', '环高架', '环路',
                        # 城市环线和快速路
                        '城市环线', '内环', '中环', '外环', '绕城环线', '城市快速环线',
                        '一环', '二环', '三环', '四环', '五环', '六环', '七环', '八环',
                        # 北京高架系统
                        '北京二环', '北京三环', '北京四环', '北京五环', '北京六环',
                        '西直门立交', '东直门立交', '北苑立交', '三元桥', '四元桥',
                        '五方桥', '六里桥', '八宝山立交', '万泉河立交', '莲花桥',
                        '长安街高架', '阜石路高架', '西三环高架', '东三环高架',
                        # 上海高架系统
                        '上海内环', '上海中环', '上海外环', '上海郊环', '延安高架', '南北高架',
                        '沪闵高架', '逸仙高架', '沪嘉高架', '鲁班高架', '中山高架',
                        '南浦大桥', '杨浦大桥', '徐浦大桥', '卢浦大桥', '黄浦江越江隧道',
                        '打浦路高架', '龙耀路高架', '陆家嘴环路', '浦东南路隧道',
                        # 广州高架系统
                        '广州内环', '广州东环', '广州北环', '广园快速', '华南快速',
                        '新港东路高架', '黄埔大道高架', '广州大道高架', '南沙港快速',
                        '琶洲大桥', '猎德大桥', '海印大桥', '江湾大桥', '解放大桥',
                        # 深圳高架系统
                        '深圳北环', '深圳南环', '滨海大道高架', '深南大道高架',
                        '皇岗路高架', '红荔路高架', '深圳湾大桥', '深港西部通道',
                        # 成都高架系统
                        '成都一环', '成都二环', '成都三环', '成都绕城高架',
                        '人民南路高架', '科华立交', '双庆立交', '红星立交',
                        '成温邛高架', '成彭高架', '成洛大道高架',
                        # 重庆高架系统
                        '重庆内环', '重庆中环', '重庆外环', '渝澳大桥', '菜园坝大桥',
                        '嘉陵江大桥', '长江大桥', '千厮门大桥', '东水门大桥',
                        '石板坡长江大桥', '朝天门长江大桥', '黄花园大桥',
                        # 武汉高架系统
                        '武汉内环', '武汉二环', '武汉三环', '武汉四环',
                        '长江一桥', '长江二桥', '长江三桥', '长江四桥', '长江五桥',
                        '汉阳大道高架', '武昌友谊大道高架', '汉口解放大道高架',
                        # 南京高架系统
                        '南京内环', '南京中环', '南京外环', '南京绕城',
                        '长江大桥', '长江二桥', '长江三桥', '长江四桥', '长江五桥',
                        '南京长江隧道', '江东路高架', '应天大街高架',
                        # 杭州高架系统
                        '杭州绕城', '杭州钱江一桥', '杭州钱江二桥', '杭州钱江三桥',
                        '杭州钱江四桥', '文晖高架', '秋石高架', '石桥路高架',
                        # 西安高架系统
                        '西安二环', '西安三环', '西安绕城', '西安北辰立交',
                        '西安城东立交', '西安城西立交', '西安城南立交',
                        # 天津高架系统
                        '天津外环', '天津中环', '天津内环', '解放南路高架',
                        '海河大桥', '解放桥', '金钟桥', '天津大桥',
                        # 其他城市高架
                        '长沙绕城高架', '长沙湘江大桥', '长沙湘府路高架',
                        '郑州东三环', '郑州西三环', '郑州北三环', '郑州南三环',
                        '青岛胶州湾高架', '青岛海湾大桥', '青岛胶州湾隧道',
                        '宁波环城高架', '宁波东环高架', '宁波西环高架',
                        '苏州绕城高架', '苏州金鸡湖大桥', '苏州独墅湖大桥',
                        # 其他高架设施
                        '跨海通道', '越江通道', '过江通道', '高架道', '高架快速', '城市高架网',
                        '立交桥系统', '互通立交', 'Y型立交', '苜蓿叶立交', '全互通立交',
                        '枢纽立交', '单喇叭立交', '双喇叭立交', '蝶式立交', '菱形立交',
                        '钻石型立交', '涡轮式立交', '环形立交', '十字立交'
                    ]
                    
                    # 用于存储每个step对应的点数范围
                    step_point_ranges = []
                    current_point_index = 0
                    
                    for i, step in enumerate(steps):
                        polyline = step['polyline'].split(';')
                        point_count = len(polyline)
                        
                        # 记录该step的点数范围
                        start_idx = current_point_index
                        end_idx = start_idx + point_count
                        step_point_ranges.append((start_idx, end_idx))
                        current_point_index = end_idx
                        
                        # 确定道路类型
                        road_type = '0'  # 默认为普通道路
                        
                        # 尝试获取road_type字段（V5 API）
                        if 'road_type' in step:
                            road_type = str(step.get('road_type', '0')).strip()
                        # 尝试从highway字段确定是否高速（V3 API）
                        elif 'highway' in step and step['highway'] == '1':
                            road_type = '1'  # 高速公路
                        
                        # 根据道路名称判断道路类型
                        road_name = step.get('road_name', '')
                        
                        # 如果road_type不是1或2，尝试根据道路名称判断
                        if road_type not in ['1', '2']:
                            # 检查是否为高速公路
                            for keyword in highway_keywords:
                                if keyword in road_name:
                                    road_type = '1'  # 高速公路
                                    print(f"根据关键词'{keyword}'判断'{road_name}'为高速公路")
                                    break
                            
                            # 如果不是高速公路，检查是否为高架路
                            if road_type == '0':
                                for keyword in elevated_keywords:
                                    if keyword in road_name:
                                        road_type = '2'  # 城市高架
                                        print(f"根据关键词'{keyword}'判断'{road_name}'为高架路")
                                        break
                            
                            # 如果是国道或省道，也标记为高速
                            if road_name.startswith('G') or road_name.startswith('S'):
                                if len(road_name) > 1 and road_name[1].isdigit():
                                    road_type = '1'  # 高速公路
                                    print(f"根据编号'{road_name}'判断为高速公路")
                        
                        # 调试：输出每个step的道路类型和名称
                        highway_info = f", highway={step.get('highway', 'N/A')}" if 'highway' in step else ""
                        print(f"路段{i+1}: 类型='{road_type}'{highway_info}, 名称={road_name}, 点数={len(polyline)}")
                        
                        # 为这个路段的所有点分配道路类型和名称
                        for point in polyline:
                            lon, lat = map(float, point.split(','))
                            # 转换坐标到WGS84
                            #lon, lat = self.gcj02_to_wgs84(lon, lat)
                            #lon, lat = self.(lon, lat)
                            coordinates.append((lon, lat))
                            road_types.append(road_type)
                            road_names.append(road_name)  # 为每个坐标点记录道路名称
                    
                    # 手动检测高速公路和高架路
                    # 如果道路名称包含相关关键词，则将其标记为相应类型
                    for i, step in enumerate(steps):
                        road_name = step.get('road_name', '')
                        road_type = '0'
                        
                        # 检查是否为高速公路
                        for keyword in highway_keywords:
                            if keyword in road_name:
                                road_type = '1'  # 高速公路
                                break
                        
                        # 如果不是高速公路，检查是否为高架路
                        if road_type == '0':
                            for keyword in elevated_keywords:
                                if keyword in road_name:
                                    road_type = '2'  # 城市高架
                                    break
                        
                        # 如果识别出特殊道路类型，更新对应点的道路类型
                        if road_type != '0':
                            # 获取这个step对应的点范围
                            if i < len(step_point_ranges):
                                start_idx, end_idx = step_point_ranges[i]
                                # 更新这些点的道路类型
                                for k in range(start_idx, end_idx):
                                    if k < len(road_types):
                                        road_types[k] = road_type
                    
                    # 调试：检查road_types数组
                    unique_types = set(road_types)
                    print(f"\n坐标点道路类型统计 (总计{len(road_types)}个点):")
                    for rt in unique_types:
                        count = road_types.count(rt)
                        print(f"  类型 '{rt}': {count}个点 ({count/len(road_types)*100:.1f}%)")
                    
                    # 确保所有road_types都是字符串
                    road_types = [str(rt).strip() for rt in road_types]
                    
                    return coordinates, road_types, road_names
                elif data['infocode'] == '10021':  # 配额超限
                    if attempt < max_retries - 1:  # 如果不是最后一次尝试
                        time.sleep(retry_delay * (attempt + 1))  # 指数退避
                        continue
                    else:
                        raise Exception(f"API配额超限，请稍后再试或更换API Key: {data}")
                else:
                    if attempt < max_retries - 1:
                        time.sleep(retry_delay * (attempt + 1))
                        continue
                    else:
                        raise Exception(f"无法获取路线，请检查输入参数或API Key: {data}")
            except Exception as e:
                if attempt < max_retries - 1:
                    time.sleep(retry_delay * (attempt + 1))
                else:
                    raise e

class RouteGenerator(QThread):
    """生成HTML地图的线程"""
    progress_updated = pyqtSignal(int)
    generation_finished = pyqtSignal(str)
    error_occurred = pyqtSignal(str)
    
    def __init__(self, excel_files, output_dir):
        super().__init__()
        self.excel_files = excel_files
        self.output_dir = output_dir
    
    def run(self):
        try:
            self.progress_updated.emit(0)
            
            # 加载所有Excel文件
            routes = []
            total_files = len(self.excel_files)
            
            for i, file_path in enumerate(self.excel_files):
                progress = int((i / total_files) * 100)
                self.progress_updated.emit(progress)
                
                try:
                    # 读取Excel文件
                    xl = pd.ExcelFile(file_path)
                    if "所有坐标点" not in xl.sheet_names:
                        continue
                        
                    df = pd.read_excel(file_path, sheet_name="所有坐标点")
                    
                    # 获取经纬度列
                    if '经度' in df.columns:
                        lon_col = '经度'
                    elif 'B' in df.columns:
                        lon_col = 'B'
                    else:
                        lon_col = 1
                        
                    if '纬度' in df.columns:
                        lat_col = '纬度'
                    elif 'C' in df.columns:
                        lat_col = 'C'
                    else:
                        lat_col = 2
                    
                    # 跳过第一行（标题行）如果没有列名为"经度"和"纬度"
                    if '经度' not in df.columns and '纬度' not in df.columns:
                        df = df.iloc[1:]
                    
                    # 创建点列表
                    point_list = []
                    for idx, row in df.iterrows():
                        try:
                            lon = float(row[lon_col])
                            lat = float(row[lat_col])
                            
                            point = {
                                'lon': lon,
                                'lat': lat,
                                # 将WGS84坐标转换为GCJ-02坐标系
                                #'lon': gcj02_from_wgs84(lon, lat)[0],
                                #'lat': gcj02_from_wgs84(lon, lat)[1],
                                'name': f'点{idx}',
                                'address': ''
                            }
                            point_list.append(point)
                        except (ValueError, TypeError):
                            continue
                    
                    if not point_list:
                        continue
                    
                    # 检查是否有道路类型信息
                    road_types = []
                    road_names = []  # 新增：存储道路名称
                    
                    if "道路类型" in df.columns:
                        print(f"文件 {file_path} 包含道路类型列")
                        
                        # 调试：打印道路类型列的唯一值
                        unique_types = df["道路类型"].unique()
                        print(f"道路类型列的唯一值: {unique_types}")
                        
                        # 打印前10个道路类型值的实际内容和类型
                        print("前10个道路类型值:")
                        for j, val in enumerate(df["道路类型"].iloc[:10]):
                            print(f"  值{j}: '{val}' (类型: {type(val).__name__})")
                        
                        for _, row in df.iterrows():
                            road_type_value = row["道路类型"]
                            
                            # 处理NaN或None值
                            if pd.isna(road_type_value) or road_type_value is None:
                                road_types.append("0")
                                continue
                            
                            # 将值转换为字符串并去除空白
                            road_type = str(road_type_value).strip()
                            
                            # 检查道路类型值
                            if road_type in ["高速公路", "1"]:
                                road_types.append("1")
                            elif road_type in ["城市高架", "2"]:
                                road_types.append("2")
                            else:
                                road_types.append("0")
                    else:
                        print(f"文件 {file_path} 不包含道路类型列")
                    
                    # 检查是否有道路名称信息
                    if "道路名称" in df.columns:
                        print(f"文件 {file_path} 包含道路名称列")
                        for _, row in df.iterrows():
                            road_name_value = row["道路名称"]
                            # 处理NaN或None值
                            if pd.isna(road_name_value) or road_name_value is None:
                                road_names.append("")
                            else:
                                road_names.append(str(road_name_value).strip())
                    else:
                        print(f"文件 {file_path} 不包含道路名称列")
                    
                    # 创建路线对象
                    route_data = {
                        'routeName': os.path.basename(file_path).replace('.xlsx', '').replace('.xls', ''),
                        'pointList': point_list
                    }
                    
                    # 如果有道路类型信息，添加到路线数据中
                    if road_types and len(road_types) == len(point_list):
                        # 确保所有道路类型都是字符串
                        road_types = [str(rt).strip() for rt in road_types]
                        route_data['road_types'] = road_types
                        print(f"成功添加 {len(road_types)} 个道路类型信息到路线 {route_data['routeName']}")
                        
                        # 调试：统计各类型道路数量
                        type_counts = {}
                        for rt in road_types:
                            type_counts[rt] = type_counts.get(rt, 0) + 1
                        for rt, count in type_counts.items():
                            print(f"  类型 '{rt}': {count}个点")
                    else:
                        print(f"警告：道路类型数量({len(road_types)})与点数量({len(point_list)})不匹配，无法添加道路类型信息")
                    
                    # 如果有道路名称信息，添加到路线数据中
                    if road_names and len(road_names) == len(point_list):
                        route_data['road_names'] = road_names
                        print(f"成功添加 {len(road_names)} 个道路名称信息到路线 {route_data['routeName']}")
                    else:
                        print(f"警告：道路名称数量({len(road_names)})与点数量({len(point_list)})不匹配，无法添加道路名称信息")
                    
                    routes.append(route_data)
                
                except Exception as e:
                    self.error_occurred.emit(f"处理文件 {file_path} 时出错: {str(e)}")
                    continue
            
            if not routes:
                raise Exception("没有可用的路线数据")
            
            # 创建地图
            self.progress_updated.emit(90)
            route_map = self.create_route_map(routes)
            
            # 保存HTML文件
            html_path = os.path.join(self.output_dir, "路线全览.html")
            route_map.save(html_path)
            # 自动打开生成的HTML文件
            webbrowser.open(f'file://{os.path.abspath(html_path)}')
            logger.info(f'地图已生成并自动打开: {html_path}')
            
            self.progress_updated.emit(100)
            self.generation_finished.emit(html_path)
            
        except Exception as e:
            self.error_occurred.emit(str(e))
    
    def create_route_map(self, routes):
        """创建包含所有路线的地图"""
        # 以第一条路线的第一个点为中心
        start_point = routes[0]['pointList'][0]
        #m = folium.Map(location=[start_point['lat'], start_point['lon']], zoom_start=10)
        m = folium.Map(
            location=[start_point['lat'], start_point['lon']], zoom_start=10,
            tiles='https://webrd03.is.autonavi.com/appmaptile?lang=zh_cn&size=1&scale=1&style=8&x={x}&y={y}&z={z}',
            attr='© <a href="https://ditu.amap.com/">高德地图</a>',
            control_scale=True   
            # self.map = folium.Map(
            # location=[center_lat, center_lon], 
            # zoom_start=10,
            # tiles='https://webrd03.is.autonavi.com/appmaptile?lang=zh_cn&size=1&scale=1&style=8&x={x}&y={y}&z={z}',
            # attr='© <a href="https://ditu.amap.com/">高德地图</a>',
            # control_scale=True
        )
        # 创建一个单独的图层组来存放所有标记（用于统一控制所有标记）
        all_markers_group = folium.FeatureGroup(name="所有标记点")
        
        # 计算所有路线的总里程
        total_distance = 0
        total_highway_distance = 0  # 高速公路总里程
        total_elevated_distance = 0  # 高架路总里程
        
        # 创建图例HTML
        legend_html = '''
        <div style="position: fixed; 
            bottom: 50px; right: 50px; width: 350px; height: auto; 
            background-color: white; border:2px solid grey; z-index:9999; 
            font-size:14px; padding: 10px; border-radius: 5px; max-height: 500px; overflow-y: auto;">
            <div style="text-align: center; font-weight: bold; margin-bottom: 5px;">路线图例</div>
        '''

        # 为每条路线创建一个特征组
        for i, route in enumerate(routes):
            color = COLORS[i % len(COLORS)]
            route_id = f"route_{i}"
            route_name = route.get('routeName', '未命名路线')
            
            # 创建一个特征组，用于控制路线的显示/隐藏
            fg = folium.FeatureGroup(name=f"{route_name}")
            
            # 检查路线数据结构
            if 'pointList' not in route or not isinstance(route['pointList'], list):
                continue

            # 添加路线标记点
            locations = []
            points = []  # 存储点对象，包含经纬度
            for point in route['pointList']:
                # 检查点数据结构
                if not isinstance(point, dict):
                    continue
                    
                if 'lat' not in point or 'lon' not in point:
                    continue
                
                # 获取点信息
                lat = point.get('lat')
                lon = point.get('lon')
                
                # 确保经纬度是数值
                try:
                    lat = float(lat)
                    lon = float(lon)
                    point['lat'] = lat  # 确保是浮点数
                    point['lon'] = lon  # 确保是浮点数
                    points.append(point)  # 添加到点列表
                    locations.append([lat, lon])
                except (ValueError, TypeError):
                    continue

            # 只有当有足够的点时才添加路线连线
            if len(locations) >= 2:
                # 使用AntPath插件绘制带方向动画的路线
                ant_path = AntPath(
                    locations,
                    color=color,
                    weight=4.0,
                    opacity=0.8,
                    tooltip=route_name,
                    delay=20000,
                    dash_array=[100, 200],
                    pulse_color=color
                )
                
                # 将路线添加到对应的特征组
                ant_path.add_to(fg)
                
                # 添加起点和终点标记到路线的特征组
                start_point = points[0]
                end_point = points[-1]
                
                # 起点标记 - 使用亮绿色标记
                folium.Marker(
                    [start_point['lat'], start_point['lon']],
                    tooltip=f"{route_name} - 起点",
                    icon=folium.Icon(color='lightgreen', icon='play', prefix='fa'),
                    popup=f"<b>{route_name}</b><br>起点"
                ).add_to(fg)  # 添加到路线图层，与路线一起显示/隐藏
                
                # 终点标记 - 使用亮红色标记
                folium.Marker(
                    [end_point['lat'], end_point['lon']],
                    tooltip=f"{route_name} - 终点",
                    icon=folium.Icon(color='darkred', icon='stop', prefix='fa'),
                    popup=f"<b>{route_name}</b><br>终点"
                ).add_to(fg)  # 添加到路线图层，与路线一起显示/隐藏
                
                # 添加方向箭头标记
                # 在路线上每隔一定距离添加一个箭头，显示行驶方向
                if len(locations) > 10:
                    # 如果点数较多，增加箭头数量，每隔约5%的点添加一个方向箭头
                    step = max(1, len(locations) // 20)
                    for j in range(step, len(locations) - step, step):
                        # 计算方向角度 - 使用前后点确定方向
                        p1 = locations[j-1]  # 前一个点
                        p2 = locations[j+1]  # 后一个点
                        
                        # 计算方向角度（弧度）
                        dx = p2[1] - p1[1]  # 经度差
                        dy = p2[0] - p1[0]  # 纬度差
                        
                        # 根据方向选择合适的图标
                        icon_name = 'arrow-right'  # 默认向右
                        
                        # 根据方向角度选择图标
                        if abs(dx) > abs(dy):  # 主要是水平方向
                            icon_name = 'arrow-right' if dx > 0 else 'arrow-left'
                        else:  # 主要是垂直方向
                            icon_name = 'arrow-up' if dy > 0 else 'arrow-down'
                        
                        # 创建自定义图标，使用更大更明显的箭头
                        arrow_icon = folium.features.DivIcon(
                            icon_size=(20, 20),
                            icon_anchor=(10, 10),
                            html=f'<div style="font-size: 18px; color:{color}; text-shadow: 0px 0px 3px white, 0px 0px 5px white; font-weight: bold;"><i class="fa fa-{icon_name}"></i></div>',
                        )
                        
                        # 创建带有自定义样式的箭头标记
                        folium.Marker(
                            locations[j],
                            icon=arrow_icon,
                            tooltip=f"{route_name} - 行驶方向"
                        ).add_to(fg)
                    
                    # 在路线中间位置添加一个更大的方向指示器
                    # mid_point_idx = len(locations) // 2
                    # mid_point = locations[mid_point_idx]
                    
                    # # 计算中点前后的方向
                    # p_before = locations[max(0, mid_point_idx - 3)]
                    # p_after = locations[min(len(locations) - 1, mid_point_idx + 3)]
                    
                    # # 计算方向
                    # dx = p_after[1] - p_before[1]  # 经度差
                    # dy = p_after[0] - p_before[0]  # 纬度差
                    
                    # # 根据方向选择合适的图标
                    # mid_icon_name = 'arrow-right'  # 默认向右
                    
                    # if abs(dx) > abs(dy):  # 主要是水平方向
                    #     mid_icon_name = 'arrow-right' if dx > 0 else 'arrow-left'
                    # else:  # 主要是垂直方向
                    #     mid_icon_name = 'arrow-up' if dy > 0 else 'arrow-down'
                    
                    # # 创建更大的中点方向指示器
                    # mid_arrow_icon = folium.features.DivIcon(
                    #     icon_size=(40, 40),
                    #     icon_anchor=(20, 20),
                    #     html=f'''
                    #         <div style="
                    #             font-size: 28px; 
                    #             color: {color}; 
                    #             text-shadow: 0px 0px 4px white, 0px 0px 6px white, 0px 0px 8px white; 
                    #             font-weight: bold;
                    #             background-color: rgba(255,255,255,0.6);
                    #             border-radius: 50%;
                    #             width: 36px;
                    #             height: 36px;
                    #             display: flex;
                    #             align-items: center;
                    #             justify-content: center;
                    #             box-shadow: 0px 0px 5px rgba(0,0,0,0.3);
                    #         ">
                    #             <i class="fa fa-{mid_icon_name}"></i>
                    #         </div>
                    #     ''',
                    # )
                    
                    # # 添加中点方向指示器
                    # folium.Marker(
                    #     mid_point,
                    #     icon=mid_arrow_icon,
                    #     tooltip=f"{route_name} - 主行驶方向"
                    # ).add_to(fg)
                
                # 将特征组添加到地图
                fg.add_to(m)
                
                # 计算路线总距离
                route_distance = self.calculate_route_distance(points)
                total_distance += route_distance
                
                # 计算高速和高架里程（如果有道路类型数据）
                highway_distance = 0
                elevated_distance = 0
                
                # 调试：打印道路类型信息
                if 'road_types' in route:
                    print(f"路线 {route_name} 的道路类型信息:")
                    road_type_counts = {}
                    for rt in route['road_types']:
                        road_type_counts[rt] = road_type_counts.get(rt, 0) + 1
                    for rt, count in road_type_counts.items():
                        print(f"  类型 '{rt}': {count}个点")
                    
                    # 检查道路类型和点数量是否匹配
                    if len(route['road_types']) != len(points):
                        print(f"警告: 道路类型数量({len(route['road_types'])})与点数量({len(points)})不匹配!")
                        # 如果道路类型数量少于点数量，使用最后一个道路类型填充
                        if len(route['road_types']) < len(points):
                            last_type = route['road_types'][-1] if route['road_types'] else '0'
                            route['road_types'].extend([last_type] * (len(points) - len(route['road_types'])))
                            print(f"已填充道路类型至{len(route['road_types'])}个点")
                
                # 调试：打印道路名称信息
                if 'road_names' in route:
                    print(f"路线 {route_name} 包含道路名称信息")
                    # 检查道路名称和点数量是否匹配
                    if len(route['road_names']) != len(points):
                        print(f"警告: 道路名称数量({len(route['road_names'])})与点数量({len(points)})不匹配!")
                
                if 'road_types' in route and len(route['road_types']) >= len(points):
                    # 计算高速和高架里程
                    segment_distances = []  # 存储每段距离
                    segment_types = []      # 存储每段类型
                    segment_names = []      # 存储每段道路名称
                    
                    for j in range(len(points) - 1):
                        if j < len(route['road_types']):
                            road_type = str(route['road_types'][j]).strip()
                            road_name = ""
                            if 'road_names' in route and j < len(route['road_names']):
                                road_name = route['road_names'][j]
                            
                            point1 = points[j]
                            point2 = points[j + 1]
                            
                            # 计算两点间距离
                            distance = self.calculate_distance(
                                point1['lat'], point1['lon'], 
                                point2['lat'], point2['lon']
                            ) / 1000  # 转换为公里
                            
                            segment_distances.append(distance)
                            segment_types.append(road_type)
                            segment_names.append(road_name)
                            
                            # 调试：打印当前点的道路类型和距离
                            if j % 100 == 0:  # 每100个点打印一次
                                print(f"点{j} -> 点{j+1}, 类型: '{road_type}', 名称: '{road_name}', 距离: {distance:.3f}公里")
                            
                            # 根据道路类型累加里程
                            if road_type == '1':  # 高速公路
                                highway_distance += distance
                            elif road_type == '2':  # 城市高架路
                                elevated_distance += distance
                    
                    # 累加到总里程
                    total_highway_distance += highway_distance
                    total_elevated_distance += elevated_distance
                    
                    # 调试：打印里程统计
                    print(f"路线 {route_name} 里程统计:")
                    print(f"  总里程: {route_distance:.2f}公里")
                    print(f"  高速里程: {highway_distance:.2f}公里 ({highway_distance/route_distance*100:.1f}%)")
                    print(f"  高架里程: {elevated_distance:.2f}公里 ({elevated_distance/route_distance*100:.1f}%)")
                    print(f"  普通道路里程: {route_distance - highway_distance - elevated_distance:.2f}公里 ({(route_distance - highway_distance - elevated_distance)/route_distance*100:.1f}%)")
                
                # 添加到图例（包含复选框和距离信息）
                legend_html += f'''
                <div style="display: flex; align-items: center; margin-bottom: 5px;">
                    <input type="checkbox" id="checkbox-{i}" checked 
                           onchange="toggleRoute({i}, this.checked)" 
                           style="margin-right: 5px;">
                    <div style="background-color: {color}; width: 15px; height: 15px; margin-right: 5px;"></div>
                    <label for="checkbox-{i}" style="cursor: pointer;">{route_name} ({round(route_distance, 2)} 公里)</label>
                </div>
                '''
                
                # 如果有高速和高架信息，添加到图例
                if highway_distance > 0 or elevated_distance > 0:
                    legend_html += f'''
                    <div id="route-info-{i}" class="route-info" style="margin-left: 20px; font-size: 12px; margin-bottom: 10px;">
                        <div>高速: {round(highway_distance, 2)} 公里 ({round(highway_distance/route_distance*100, 1)}%)</div>
                        <div>高架: {round(elevated_distance, 2)} 公里 ({round(elevated_distance/route_distance*100, 1)}%)</div>
                        <div>普通: {round(route_distance - highway_distance - elevated_distance, 2)} 公里 ({round((route_distance - highway_distance - elevated_distance)/route_distance*100, 1)}%)</div>
                    </div>
                    '''
        
        # 添加全选/全不选按钮
        legend_html += '''
        <div style="margin-top: 10px; border-top: 1px solid #ccc; padding-top: 10px; text-align: center;">
            <button onclick="toggleAllRoutes(true)" style="margin-right: 10px; padding: 3px 8px;">全选</button>
            <button onclick="toggleAllRoutes(false)" style="padding: 3px 8px;">全不选</button>
        </div>
        '''
        
        # 添加所有路线的总里程到图例
        legend_html += f'''
        <div style="border-top: 1px solid #ccc; margin-top: 5px; padding-top: 5px;">
            <div style="font-weight: bold; text-align: center;">总里程: {round(total_distance, 2)} 公里</div>
            <div style="text-align: center;">高速总里程: {round(total_highway_distance, 2)} 公里 ({round(total_highway_distance/total_distance*100 if total_distance > 0 else 0, 1)}%)</div>
            <div style="text-align: center;">高架总里程: {round(total_elevated_distance, 2)} 公里 ({round(total_elevated_distance/total_distance*100 if total_distance > 0 else 0, 1)}%)</div>
            <div style="text-align: center;">普通道路总里程: {round(total_distance - total_highway_distance - total_elevated_distance, 2)} 公里 ({round((total_distance - total_highway_distance - total_elevated_distance)/total_distance*100 if total_distance > 0 else 0, 1)}%)</div>
        </div>
        '''
        
        # 完成图例HTML
        legend_html += '</div>'
        
        # 将图例添加到地图
        m.get_root().html.add_child(folium.Element(legend_html))

        # 添加地图说明
        map_info_html = '''
        <div id="map-info-box" style="position: fixed; 
            top: 10px; left: 50px; width: 250px; 
            background-color: white; border:2px solid grey; z-index:9999; 
            font-size:14px; padding: 10px; border-radius: 5px;">
            <div style="font-weight: bold; margin-bottom: 5px;">地图说明:</div>
            <div style="margin-bottom: 3px;"><span style="color:lightgreen;">●</span> 浅绿色标记为路线起点</div>
            <div style="margin-bottom: 3px;"><span style="color:darkred;">●</span> 深红色标记为路线终点</div>
            <div style="margin-bottom: 3px;">➤ 彩色箭头表示行驶方向</div>
            <div style="margin-bottom: 3px;">路线颜色与右侧图例对应</div>
            <div style="margin-top: 10px; text-align: center;">
                <button id="toggle-markers-btn" style="padding: 5px 10px; cursor: pointer; background-color: #f0f0f0; border: 1px solid #ccc; border-radius: 4px; font-weight: bold;">隐藏所有标记</button>
            </div>
        </div>
        '''
        
        # 将地图说明添加到地图
        m.get_root().html.add_child(folium.Element(map_info_html))
        
        # 添加隐藏标记的JavaScript代码
        marker_control_js = '''
        <script>
        // 页面加载完成后执行
        document.addEventListener('DOMContentLoaded', function() {
            // 设置按钮点击事件
            var btn = document.getElementById('toggle-markers-btn');
            if (!btn) {
                console.error('找不到标记按钮');
                return;
            }
            
            var markersHidden = false;
            var originalMarkerPane = null;
            var originalShadowPane = null;
            var originalPopupPane = null;
            
            // 点击按钮时执行
            btn.onclick = function() {
                if (markersHidden) {
                    // 如果标记已隐藏，则恢复标记（刷新页面）
                    window.location.reload();
                    return;
                }
                
                // 隐藏标记
                markersHidden = true;
                
                // 更新按钮文本
                this.textContent = '显示所有标记';
                this.style.backgroundColor = '#ffcccc';
                
                // 创建或获取样式元素
                var styleEl = document.getElementById('marker-style');
                if (!styleEl) {
                    styleEl = document.createElement('style');
                    styleEl.id = 'marker-style';
                    document.head.appendChild(styleEl);
                }
                
                // 使用极端的CSS选择器和属性覆盖来隐藏标记
                styleEl.textContent = `
                    .leaflet-marker-pane *,
                    .leaflet-shadow-pane *,
                    .leaflet-popup-pane * {
                        display: none !important;
                        visibility: hidden !important;
                        opacity: 0 !important;
                        width: 0 !important;
                        height: 0 !important;
                        overflow: hidden !important;
                        position: absolute !important;
                        left: -9999px !important;
                        top: -9999px !important;
                        pointer-events: none !important;
                        z-index: -9999 !important;
                    }
                `;
                
                // 尝试直接移除标记元素
                setTimeout(function() {
                    var markerPane = document.querySelector('.leaflet-marker-pane');
                    var shadowPane = document.querySelector('.leaflet-shadow-pane');
                    var popupPane = document.querySelector('.leaflet-popup-pane');
                    
                    if (markerPane) markerPane.innerHTML = '';
                    if (shadowPane) shadowPane.innerHTML = '';
                    if (popupPane) popupPane.innerHTML = '';
                    
                    console.log('已清空标记容器');
                }, 100);
            };
            
            console.log('标记控制功能已初始化');
        });
        </script>
        '''
        
        m.get_root().html.add_child(folium.Element(marker_control_js))
        
        # 添加JavaScript代码，用于控制路线和标记的显示/隐藏
        route_control_js = '''
        <script>
        document.addEventListener('DOMContentLoaded', function() {
            // 延迟执行，确保地图已完全加载
            setTimeout(function() {
                // 获取所有图层控制项
                var layerControls = document.querySelectorAll('.leaflet-control-layers-selector');
                var routeLayers = [];
                var layerNames = [];
                
                // 存储图层控制项和对应的名称
                for (var i = 0; i < layerControls.length; i++) {
                    routeLayers.push(layerControls[i]);
                    // 获取图层名称
                    var label = layerControls[i].nextSibling;
                    while (label && label.nodeType !== 1) {
                        label = label.nextSibling;
                    }
                    if (label) {
                        layerNames.push(label.textContent.trim());
                    } else {
                        layerNames.push('未命名图层');
                    }
                }
                
                // 定义切换路线显示/隐藏的函数
                window.toggleRoute = function(index, show) {
                    // 查找对应名称的图层控制项
                    var checkboxId = 'checkbox-' + index;
                    var checkboxLabel = document.querySelector('label[for="' + checkboxId + '"]');
                    
                    if (checkboxLabel) {
                        var routeName = checkboxLabel.textContent.trim().split(' (')[0];
                        
                        // 查找匹配的图层控制项
                        for (var i = 0; i < layerNames.length; i++) {
                            if (layerNames[i] === routeName) {
                                // 如果当前状态与目标状态不同，则点击切换
                                if (routeLayers[i].checked !== show) {
                                    routeLayers[i].click();
                                }
                                break;
                            }
                        }
                        
                        // 同步更新图例中的路线信息显示/隐藏
                        var routeInfoDiv = document.getElementById('route-info-' + index);
                        if (routeInfoDiv) {
                            routeInfoDiv.style.display = show ? 'block' : 'none';
                        }
                    }
                };
                
                // 定义全选/全不选函数
                window.toggleAllRoutes = function(show) {
                    // 更新所有复选框
                    var checkboxes = document.querySelectorAll('input[id^="checkbox-"]');
                    for (var i = 0; i < checkboxes.length; i++) {
                        checkboxes[i].checked = show;
                        var index = parseInt(checkboxes[i].id.replace('checkbox-', ''));
                        toggleRoute(index, show);
                    }
                };
                
                // 将图层控制器与自定义复选框同步
                for (var i = 0; i < routeLayers.length; i++) {
                    (function(index, layerName) {
                        routeLayers[index].addEventListener('change', function() {
                            // 查找对应名称的复选框
                            var checkboxes = document.querySelectorAll('input[id^="checkbox-"]');
                            for (var j = 0; j < checkboxes.length; j++) {
                                var checkboxId = checkboxes[j].id;
                                var checkboxLabel = document.querySelector('label[for="' + checkboxId + '"]');
                                if (checkboxLabel && checkboxLabel.textContent.trim().split(' (')[0] === layerName) {
                                    checkboxes[j].checked = this.checked;
                                    
                                    // 同步更新图例中的路线信息显示/隐藏
                                    var routeIndex = parseInt(checkboxId.replace('checkbox-', ''));
                                    var routeInfoDiv = document.getElementById('route-info-' + routeIndex);
                                    if (routeInfoDiv) {
                                        routeInfoDiv.style.display = this.checked ? 'block' : 'none';
                                    }
                                    
                                    break;
                                }
                            }
                        });
                    })(i, layerNames[i]);
                }
                
            }, 1500);  // 延迟1.5秒，确保地图已完全渲染
        });
        </script>
        '''
        
        m.get_root().html.add_child(folium.Element(route_control_js))
        
        # 添加图层控制器（折叠状态）
        folium.LayerControl(collapsed=True).add_to(m)

        return m
    
    def calculate_distance(self, lat1, lon1, lat2, lon2):
        """计算两点间的距离（单位：米）"""
        # 地球半径（米）
        R = 6371000
        
        # 将经纬度转换为弧度
        lat1_rad = math.radians(lat1)
        lon1_rad = math.radians(lon1)
        lat2_rad = math.radians(lat2)
        lon2_rad = math.radians(lon2)
        
        # 计算差值
        dlat = lat2_rad - lat1_rad
        dlon = lon2_rad - lon1_rad
        
        # Haversine公式
        a = math.sin(dlat/2)**2 + math.cos(lat1_rad) * math.cos(lat2_rad) * math.sin(dlon/2)**2
        c = 2 * math.atan2(math.sqrt(a), math.sqrt(1-a))
        distance = R * c
        
        return distance
    
    def calculate_route_distance(self, points):
        """计算路线总距离（单位：公里）"""
        if len(points) < 2:
            return 0
            
        total_distance = 0
        for i in range(len(points) - 1):
            lat1 = points[i]['lat']
            lon1 = points[i]['lon']
            lat2 = points[i+1]['lat']
            lon2 = points[i+1]['lon']
            
            total_distance += self.calculate_distance(lat1, lon1, lat2, lon2)
        
        # 转换为公里并保留两位小数
        return round(total_distance / 1000, 2)

class MainWindow(QMainWindow):
    def __init__(self):
        # 性能优化配置
        self.MAX_CONCURRENT_FILES = 10  # 最大并发文件处理数
        self.BATCH_SIZE = 5  # 批量处理大小
        self.CACHE_CLEANUP_INTERVAL = 100  # 缓存清理间隔(文件数)
        self.processed_files_count = 0  # 已处理文件计数器
        super().__init__()
        self.setWindowTitle(f"高德导航路线可视化工具 {VERSION}")
        self.setGeometry(100, 100, 1200, 800)  # 增大初始窗口尺寸
        self.setMinimumSize(1000, 700)  # 设置最小窗口尺寸
        self.waypoints = []
        self.all_coordinates = []
        self.route_segments = []
        self.excel_files = []
        self.json_files = []
        self.json_data_list = []
        self.routes_result = []
        self.calc_threads = []
        
        # 高德地图API设置
        # 主API密钥
        self.key = '5cd11205cc7744d742da10dda92daecd'
        
        # 备用API密钥列表 - 如果主密钥不起作用，可以尝试这些
        self.backup_keys = [
            '8325164e247e15eea68b59e89200988b',  # 备用密钥1
            '3fabc36268a955439fc99a589aacbd87',  # 备用密钥2
            '2b2d86f7b4f48047e7d5b3cec6e9f51f'   # 备用密钥3
        ]
        
        # 设置窗口图标
        self.setWindowIcon(QIcon(self.get_icon_path()))
        
        # 【自动路线规划系统属性】
        self.locations = []                    # 用户输入的地点名称列表
        self.coordinates = []                  # 获取到的坐标数据
        self.valid_locations = []              # 有效的地点（带坐标）
        self.map_file_path = None
        self.route_data = []                   # 生成的路线数据
        self.combined_map_path = None          # 所有路线的综合地图
        self.show_heatmap = False              # 是否显示热力图
        
        # 【路线规划配置参数】
        self.route_config = {
            'waypoint_min_distance': 2,        # 最小距离：2公里
            'waypoint_max_distance': 15,       # 最大距离：15公里
            'between_waypoint_min': 1.5,       # 途径点间最小距离
            'between_waypoint_max': 12,        # 途径点间最大距离
            'similarity_threshold': 0.6,       # 路线相似度阈值
            'enable_deduplication': True,      # 启用去重功能
        }
        
        self.init_ui()
        
        # 设置全局样式
        self.setStyleSheet("""
            QMainWindow { background-color: #f7f9fa; }
            QGroupBox { border: 1px solid #b0bec5; border-radius: 6px; margin-top: 10px; padding-top: 10px; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 3px; font-weight: bold; color: #37474f; }
            QPushButton { background-color: #1976d2; color: white; border: none; border-radius: 6px; padding: 10px 20px; font-size: 20px; }
            QPushButton:hover { background-color: #1565c0; }
            QPushButton:disabled { background-color: #b0bec5; color: #eceff1; }
            QProgressBar { border: 1px solid #b0bec5; border-radius: 3px; text-align: center; height: 18px; }
            QProgressBar::chunk { background-color: #1976d2; }
            QListWidget, QTextEdit, QLineEdit { border: 1px solid #b0bec5; border-radius: 3px; padding: 4px; font-size: 20px; }
            QLabel { font-size: 20px; color: #263238; }
            QTabWidget::pane { border: 1px solid #b0bec5; border-radius: 6px; }
            QTabBar::tab { background: #e3f2fd; border: 1px solid #b0bec5; border-bottom: none; border-top-left-radius: 6px; border-top-right-radius: 6px; padding: 8px 20px; font-size: 25px; }
            QTabBar::tab:selected { background: #ffffff; color: #1976d2; font-weight: bold; }
        """)
    
    def get_icon_path(self):
        """获取图标路径"""
        # 这里可以添加逻辑来查找图标文件
        # 如果没有找到，返回None
        return None
    
    def init_ui(self):
        # 创建主布局
        main_widget = QWidget()
        main_layout = QVBoxLayout(main_widget)
        
        # 顶部横向布局（左侧标题，右侧关于按钮）
        top_layout = QHBoxLayout()
        title_label = QLabel("高德地图路线规划工具")
        title_label.setStyleSheet("font-size: 25px; font-weight: bold; color: #1976d2;")
        top_layout.addWidget(title_label)
        top_layout.addStretch(1)
        self.about_btn = QPushButton("关于")
        self.about_btn.setFixedWidth(100)
        self.about_btn.setFixedHeight(52)
        self.about_btn.clicked.connect(self.show_about_dialog)
        top_layout.addWidget(self.about_btn)
        main_layout.addLayout(top_layout)
        
        # 创建选项卡
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)
        
        # 创建"自动路线规划"选项卡
        self.create_auto_route_tab()
        
        # 创建"一键处理"选项卡
        self.create_one_click_tab()
        
        # 创建"路线生成"选项卡
        self.create_route_tab()
        
        # 创建"地图生成"选项卡
        self.create_map_tab()
        
        # 设置中心窗口部件
        self.setCentralWidget(main_widget)
        # 启动时刷新生成路线按钮的状态（基于当前有效坐标数）
        try:
            # 方法可能在类中稍后定义，但运行时可以访问
            self.refresh_generate_button_state()
        except Exception:
            pass
    
    def create_route_tab(self):
        """创建路线生成选项卡（容错：确保即使发生异常也能创建占位Tab）"""
        try:
            route_tab = QWidget()
            route_layout = QVBoxLayout(route_tab)
            
            # API Key组
            key_group = QGroupBox("高德地图API设置")
            key_layout = QHBoxLayout()
            key_label = QLabel("API Key:")
            self.key_input = QLineEdit(self.key)
            key_layout.addWidget(key_label)
            key_layout.addWidget(self.key_input)
            key_group.setLayout(key_layout)
            route_layout.addWidget(key_group)
            
            # 导入JSON组
            import_group = QGroupBox("导入JSON文件")
            import_layout = QVBoxLayout()
            
            self.import_btn = QPushButton("导入JSON文件")
            self.import_btn.clicked.connect(self.import_json)
            self.import_btn.setFixedHeight(52)
            import_layout.addWidget(self.import_btn)
            
            self.import_status = QLabel("请导入JSON文件...")
            import_layout.addWidget(self.import_status)
            
            self.route_list = QListWidget()
            import_layout.addWidget(self.route_list)
            
            # 新增：移除所选和清空列表按钮（等宽）
            json_btn_layout = QHBoxLayout()
            json_btn_layout.setSpacing(12)
            json_btn_layout.setContentsMargins(0, 0, 0, 0)
            self.remove_json_btn = QPushButton("移除所选")
            self.remove_json_btn.clicked.connect(self.remove_selected_json)
            self.remove_json_btn.setFixedHeight(52)
            self.remove_json_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            self.clear_json_btn = QPushButton("清空列表")
            self.clear_json_btn.clicked.connect(self.clear_json_list)
            self.clear_json_btn.setFixedHeight(52)
            self.clear_json_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            json_btn_layout.addWidget(self.remove_json_btn, 1)
            json_btn_layout.addWidget(self.clear_json_btn, 1)
            import_layout.addLayout(json_btn_layout)
            # 结束新增
            
            import_group.setLayout(import_layout)
            route_layout.addWidget(import_group)
            
            # 进度条
            self.progress_bar = QProgressBar()
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(0)
            route_layout.addWidget(self.progress_bar)
            
            # 操作按钮组
            button_group = QGroupBox("操作")
            button_layout = QHBoxLayout()
            button_layout.setSpacing(12)
            button_layout.setContentsMargins(0, 0, 0, 0)
            
            self.calculate_btn = QPushButton("计算路线")
            self.calculate_btn.clicked.connect(self.calculate_route)
            self.calculate_btn.setEnabled(False)
            self.calculate_btn.setFixedHeight(52)
            self.calculate_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            
            self.export_btn = QPushButton("导出所有文件")
            self.export_btn.clicked.connect(self.export_all)
            self.export_btn.setEnabled(False)
            self.export_btn.setFixedHeight(52)
            self.export_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            
            button_layout.addWidget(self.calculate_btn, 1)
            button_layout.addWidget(self.export_btn, 1)
            
            button_group.setLayout(button_layout)
            route_layout.addWidget(button_group)
            
            # 日志组
            log_group = QGroupBox("日志")
            log_layout = QVBoxLayout()
            
            self.log_text = QTextEdit()
            self.log_text.setReadOnly(True)
            log_layout.addWidget(self.log_text)
            
            log_group.setLayout(log_layout)
            route_layout.addWidget(log_group)
        except Exception as e:
            logger.error(f"create_route_tab 初始化失败: {str(e)}", exc_info=True)
            # 创建占位Tab以避免NameError
            route_tab = QWidget()
            placeholder_layout = QVBoxLayout(route_tab)
            placeholder_label = QLabel(f"路线选项卡初始化失败: {str(e)}")
            placeholder_layout.addWidget(placeholder_label)
        finally:
            # 确保无论是否发生异常，都能向tab_widget添加Tab
            try:
                self.tab_widget.addTab(route_tab, "路线生成")
            except Exception as e:
                logger.error(f"向tab_widget添加路线选项卡失败: {str(e)}", exc_info=True)
                # 无法添加Tab时，确保程序不崩溃
                pass
    
    def create_map_tab(self):
        """创建地图生成选项卡"""
        map_tab = QWidget()
        map_layout = QVBoxLayout(map_tab)
        
        # 文件选择组
        file_group = QGroupBox("选择Excel文件")
        file_layout = QVBoxLayout()
        
        self.select_btn = QPushButton("添加Excel文件")
        self.select_btn.clicked.connect(self.select_excel_files)
        self.select_btn.setFixedHeight(52)
        file_layout.addWidget(self.select_btn)
        
        self.file_list = QListWidget()
        file_layout.addWidget(self.file_list)
        
        # 文件操作按钮（等宽）
        file_button_layout = QHBoxLayout()
        file_button_layout.setSpacing(12)
        file_button_layout.setContentsMargins(0, 0, 0, 0)
        
        self.remove_btn = QPushButton("移除所选")
        self.remove_btn.clicked.connect(self.remove_selected_files)
        self.remove_btn.setFixedHeight(52)
        self.remove_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        
        self.clear_btn = QPushButton("清空列表")
        self.clear_btn.clicked.connect(self.clear_files)
        self.clear_btn.setFixedHeight(52)
        self.clear_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        
        file_button_layout.addWidget(self.remove_btn, 1)
        file_button_layout.addWidget(self.clear_btn, 1)
        
        file_layout.addLayout(file_button_layout)
        file_group.setLayout(file_layout)
        map_layout.addWidget(file_group)
        
        # 生成地图组
        generate_group = QGroupBox("生成地图")
        generate_layout = QVBoxLayout()
        
        self.generate_btn = QPushButton("生成地图")
        self.generate_btn.clicked.connect(self.generate_map)
        self.generate_btn.setFixedHeight(52)
        generate_layout.addWidget(self.generate_btn)
        
        self.map_progress = QProgressBar()
        self.map_progress.setRange(0, 100)
        self.map_progress.setValue(0)
        generate_layout.addWidget(self.map_progress)
        
        generate_group.setLayout(generate_layout)
        map_layout.addWidget(generate_group)
        
        # 地图日志组
        map_log_group = QGroupBox("日志")
        map_log_layout = QVBoxLayout()
        
        self.map_log = QTextEdit()
        self.map_log.setReadOnly(True)
        map_log_layout.addWidget(self.map_log)
        
        map_log_group.setLayout(map_log_layout)
        map_layout.addWidget(map_log_group)
        
        self.tab_widget.addTab(map_tab, "地图生成")
    
    def create_auto_route_tab(self):
        """创建自动路线规划选项卡，采用左右分栏布局"""
        auto_tab = QWidget()
        
        # 创建主布局为水平布局，替代分割器，使宽度不可调节
        main_layout = QHBoxLayout(auto_tab)
        
        # ========== 左侧面板：地点输入和查询 ==========
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        
        # 地点输入和查询组
        control_frame = QGroupBox("📍 地点输入和查询")
        control_layout = QGridLayout(control_frame)
        control_layout.setSpacing(15)  # 增加组件间距
        control_layout.setContentsMargins(15, 15, 15, 15)  # 增加内边距
        
        # 城市选择（独占一行）
        # 城市标签 - 加大字体
        city_label = QLabel("城市:")
        font = city_label.font()
        font.setPointSize(20)  # 字体大小设为12pt
        font.setBold(True)
        city_label.setFont(font)
        control_layout.addWidget(city_label, 0, 0, 1, 1)  # 行0，列0，占1行1列
        
        self.city_combo = QComboBox()
        self.city_combo.setFixedWidth(220)  # 加宽城市选择框
        self.city_combo.addItem("请选择城市", "")
        for city in sorted(CITY_DISTRICTS.keys()):
            self.city_combo.addItem(city, city)
        self.city_combo.currentIndexChanged.connect(self.on_city_changed)
        control_layout.addWidget(self.city_combo, 0, 1, 1, 2)  # 行0，列1，占1行2列
        
        # 显示所有地点地图按钮（独占一行）
        self.show_map_btn = QPushButton("🗺️ 显示所有地点地图")
        self.show_map_btn.clicked.connect(self.show_all_locations_map)
        self.show_map_btn.setFixedHeight(52)  # 增加按钮高度
        self.show_map_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        control_layout.addWidget(self.show_map_btn, 1, 0, 1, 4)  # 行1，列0，占1行4列
        
        # 行政区选择
        # 行政区标签 - 加大字体
        district_label = QLabel("行政区:")
        district_label.setFont(font)  # 使用相同的大字体
        control_layout.addWidget(district_label, 2, 0, 1, 4)  # 行2，列0，占1行4列
        
        # 行政区多选框容器
        self.district_group = QWidget()
        self.district_layout = QGridLayout(self.district_group)  # 使用网格布局，分散显示
        self.district_layout.setContentsMargins(0, 0, 0, 0)
        self.district_layout.setSpacing(10)  # 增加间距
        
        self.district_checkboxes = {}
        
        # 添加全区域选项
        self.all_districts_checkbox = QCheckBox("全区域")
        self.all_districts_checkbox.setChecked(True)
        self.all_districts_checkbox.stateChanged.connect(self.on_all_districts_changed)
        self.district_layout.addWidget(self.all_districts_checkbox, 0, 0, 1, 4)  # 行0，列0，占1行4列
        
        # 将self.on_district_changed方法修改为兼容网格布局的更新
        self.district_checkboxes = {}
        
        # 初始化为空布局，后续在on_city_changed中动态添加
        
        control_layout.addWidget(self.district_group, 3, 0, 3, 4)  # 行3，列0，占4行4列（改为4列跨度）
        
        # 生活场景选择
        # 生活场景标签 - 加大字体
        scene_label = QLabel("生活场景:")
        scene_label.setFont(font)  # 使用相同的大字体
        control_layout.addWidget(scene_label, 7, 0, 1, 3)  # 行7，列0，占1行3列
        
        # 生活场景多选框容器
        self.scene_group = QWidget()
        self.scene_layout = QGridLayout(self.scene_group)  # 使用网格布局，分散显示
        self.scene_layout.setContentsMargins(0, 0, 0, 0)
        self.scene_layout.setSpacing(10)  # 增加间距
        
        # 分4-6行显示生活场景，每行2-3个
        self.scenes = ["学校", "医院", "公园", "商场", "美食街", "酒店", "加油站", "银行", "地铁站", "公交站"]
        self.scene_checkboxes = {}
        
        # 添加生活场景到网格布局
        cols = 3  # 每行3个
        for i, scene in enumerate(self.scenes):
            checkbox = QCheckBox(scene)
            checkbox.setFixedWidth(150)  # 固定宽度，确保显示完整
            # 计算行列位置
            row = i // cols
            col = i % cols
            self.scene_layout.addWidget(checkbox, row, col)
            self.scene_checkboxes[scene] = checkbox
        
        control_layout.addWidget(self.scene_group, 8, 0, 4, 3)  # 行8，列0，占4行3列
        
        # 操作按钮组（两行布局）
        # 第一行按钮
        self.search_scene_btn = QPushButton("🔍 搜索场景地点")
        self.search_scene_btn.clicked.connect(self.search_scene_locations)
        self.search_scene_btn.setFixedHeight(52)  # 增加按钮高度
        self.search_scene_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        control_layout.addWidget(self.search_scene_btn, 12, 0, 1, 3)  # 行12，列0，占1行3列
        
        # 第二行按钮
        self.delete_selected_btn = QPushButton("🗑️ 删除选中地点")
        self.delete_selected_btn.clicked.connect(self.delete_selected_locations)
        self.delete_selected_btn.setFixedHeight(52)  # 增加按钮高度
        self.delete_selected_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        control_layout.addWidget(self.delete_selected_btn, 13, 0, 1, 3)  # 行13，列0，占1行3列
        
        # 第三行按钮
        self.import_loc_btn = QPushButton("📥 导入Excel")
        self.import_loc_btn.clicked.connect(self.import_locations_from_excel)
        self.import_loc_btn.setFixedHeight(52)  # 增加按钮高度
        self.import_loc_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        control_layout.addWidget(self.import_loc_btn, 14, 0, 1, 1)  # 行14，列0，占1行1列
        
        self.fetch_btn = QPushButton("🔍 批量获取坐标")
        self.fetch_btn.clicked.connect(self.fetch_coordinates)
        self.fetch_btn.setFixedHeight(52)  # 增加按钮高度
        self.fetch_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        control_layout.addWidget(self.fetch_btn, 14, 1, 1, 2)  # 行14，列1，占1行2列
        
        left_layout.addWidget(control_frame)
        
        # 处理日志（左侧面板）
        response_frame = QGroupBox("📝 处理日志")
        response_layout = QVBoxLayout(response_frame)
        
        self.response_text = QTextEdit()
        self.response_text.setReadOnly(True)
        self.response_text.setStyleSheet("font-size: 10pt;")  # 设置合适的字体大小
        response_layout.addWidget(self.response_text)

        # 批量获取坐标状态标签（显示在处理日志下方）
        self.fetch_status_label = QLabel("")
        self.fetch_status_label.setStyleSheet("color: blue")
        response_layout.addWidget(self.fetch_status_label)
        
        left_layout.addWidget(response_frame, 1)  # 权重1，占较大空间
        
        # ========== 右侧面板：地点坐标信息、路线规划和日志 ==========
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        
        # 地点坐标信息（占更大比例）
        table_frame = QGroupBox("📊 地点坐标信息")
        table_layout = QVBoxLayout(table_frame)
        
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["序号", "地点名称", "经度", "纬度", "状态"])
        self.tree.setColumnWidth(0, 120)
        self.tree.setColumnWidth(1, 300)
        self.tree.setColumnWidth(2, 120)
        self.tree.setColumnWidth(3, 120)
        self.tree.setColumnWidth(4, 100)
        # 不设置固定高度，让其自动占满剩余空间的大部分
        table_layout.addWidget(self.tree)
        
        # 设置table_frame的拉伸系数，让其占据更多空间
        right_layout.addWidget(table_frame, 3)  # 权重3
        
        # 路线规划设置
        route_frame = QGroupBox("🛣️ 路线规划设置")
        route_layout = QHBoxLayout(route_frame)
        route_layout.setSpacing(15)
        route_layout.setContentsMargins(15, 15, 15, 15)
        
        route_layout.addWidget(QLabel("生成路线数:"))
        self.route_num_spin = QSpinBox()
        self.route_num_spin.setRange(1, 50)
        self.route_num_spin.setValue(3)
        self.route_num_spin.setFixedWidth(80)
        route_layout.addWidget(self.route_num_spin)
        
        route_layout.addWidget(QLabel("每条路线途径点数:"))
        self.waypoint_spin = QSpinBox()
        self.waypoint_spin.setRange(0, 10)
        self.waypoint_spin.setValue(2)
        self.waypoint_spin.setFixedWidth(80)
        route_layout.addWidget(self.waypoint_spin)
        
        route_layout.addStretch()
        
        self.generate_route_btn = QPushButton("🚀 自动生成测试路线")
        self.generate_route_btn.clicked.connect(self.start_generating_routes)
        self.generate_route_btn.setFixedHeight(52)  # 增加按钮高度
        self.generate_route_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        
        self.all_routes_button = QPushButton("🗺️ 查看所有路线地图")
        self.all_routes_button.clicked.connect(self.view_all_routes_map)
        self.all_routes_button.setEnabled(False)
        self.all_routes_button.setFixedHeight(52)  # 增加按钮高度
        self.all_routes_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        # 导出按钮（已移动至路线规划设置栏内）
        self.export_excel_btn = QPushButton("💾 导出到Excel")
        self.export_excel_btn.clicked.connect(self.export_excel)
        self.export_excel_btn.setFixedHeight(52)
        self.export_excel_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        
        route_layout.addWidget(self.generate_route_btn, 1)
        route_layout.addWidget(self.all_routes_button, 1)
        route_layout.addWidget(self.export_excel_btn, 1)

        right_layout.addWidget(route_frame, 1)  # 权重1
        
        # 状态栏
        status_frame = QWidget()
        status_layout = QHBoxLayout(status_frame)
        status_layout.setContentsMargins(15, 10, 15, 10)
        
        self.status_label = QLabel("就绪")
        self.status_label.setStyleSheet("color: green")
        status_layout.addWidget(self.status_label)
        
        status_layout.addStretch()
        
        right_layout.addWidget(status_frame, 0)  # 权重0，占最小空间
        
        # 设置左右面板比例：平分主界面宽度（1:1）
        main_layout.addWidget(left_widget, 1)  # 权重1，平分宽度
        main_layout.addWidget(right_widget, 1)  # 权重1，平分宽度
        
        # 更新日志
        self.update_api_response("✅ 应用启动成功！请开始输入地点名称。")
        
        self.tab_widget.addTab(auto_tab, "自动路线规划")
    
    def refresh_generate_button_state(self):
        """根据当前有效坐标数量刷新“自动生成测试路线”按钮的可用性（至少需要6个点）。"""
        try:
            valid_count = len(getattr(self, 'valid_locations', []) or [])
            if hasattr(self, 'generate_route_btn'):
                self.generate_route_btn.setEnabled(valid_count >= 6)
        except Exception:
            pass
    
    def on_city_changed(self):
        """城市选择变化时更新行政区选项"""
        # 清除所有已有的行政区多选框
        for checkbox in self.district_checkboxes.values():
            checkbox.setParent(None)
        self.district_checkboxes.clear()
        
        # 获取选中的城市
        city = self.city_combo.currentData()
        
        if city and city in CITY_DISTRICTS:
            # 获取该城市的所有行政区
            districts = CITY_DISTRICTS[city]
            
            # 添加行政区多选框到网格布局，分散显示
            cols = 4  # 每行4个（调整为4列）
            for i, district in enumerate(districts):
                checkbox = QCheckBox(district)
                checkbox.setFixedWidth(90)  # 固定宽度，确保显示完整
                checkbox.setMinimumWidth(90)
                checkbox.stateChanged.connect(self.on_district_changed)
                # 计算行列位置
                row = i // cols + 1  # 从行1开始，行0用于全区域选项
                col = i % cols
                self.district_layout.addWidget(checkbox, row, col)
                self.district_checkboxes[district] = checkbox
        
        # 重置全区域选项
        self.all_districts_checkbox.setChecked(True)
    
    def on_all_districts_changed(self, state):
        """全区域选项变化时处理其他选项的状态"""
        if state == Qt.Checked:
            # 如果全区域被选中，禁用所有其他行政区选项
            for checkbox in self.district_checkboxes.values():
                checkbox.setChecked(False)
                checkbox.setEnabled(False)
        else:
            # 如果全区域未被选中，启用所有其他行政区选项
            for checkbox in self.district_checkboxes.values():
                checkbox.setEnabled(True)
    
    def on_district_changed(self, state):
        """单个行政区选项变化时处理全区域选项的状态"""
        # 检查是否有任何行政区被选中
        any_district_checked = any(checkbox.isChecked() for checkbox in self.district_checkboxes.values())
        
        if any_district_checked:
            # 如果有任何行政区被选中，取消全区域选项
            self.all_districts_checkbox.setChecked(False)
        
    # 移除了错误的init_ui方法，直接在MainWindow的__init__方法中设置最小尺寸
    # def init_ui(self):
    #     """初始化主界面"""
    #     super().init_ui()
    #     
    #     # 设置主窗口最小尺寸
    #     self.setMinimumSize(1000, 700)
    #     
    #     # 其他初始化代码保持不变
    
    # ============= 自动路线规划系统核心方法 =============
    
    # def add_location(self):
    #     """添加单个地点"""
    #     location = self.location_entry.text().strip()
    #     if not location:
    #         QMessageBox.showwarning("警告", "请输入地点名称")
    #         return
    #     
    #     if location in self.locations:
    #         QMessageBox.showwarning("警告", "此地点已添加")
    #         return
    #     
    #     self.locations.append(location)
    #     self.location_entry.clear()
    #     
    #     # 更新表格，添加新地点
    #     self._update_location_table()
    #     
    #     self.update_api_response(f"✅ 已添加地点: {location}")
    #     self.status_label.setText(f"已添加 {len(self.locations)} 个地点")
    #     self.status_label.setStyleSheet("color: blue")
    
    def _update_location_table(self):
        """更新地点表格，确保所有地点都显示"""
        try:
            self.tree.clear()
            
            all_items = []
            
            # 1. 添加已获取坐标的地点
            for loc in self.coordinates:
                all_items.append((loc['name'], True, loc['lon'], loc['lat'], "✅ 获取成功"))
            
            # 2. 添加待查询的地点
            for location in self.locations:
                if not any(item[0] == location for item in all_items):
                    all_items.append((location, False, None, None, "⏳ 待查询"))
            
            # 3. 去重并按名称排序
            seen_names = set()
            unique_items = []
            for item in all_items:
                if item[0] not in seen_names:
                    unique_items.append(item)
                    seen_names.add(item[0])
            
            # 4. 显示到表格
            for i, (name, has_coords, lon, lat, status) in enumerate(sorted(unique_items, key=lambda x: x[0]), 1):
                if has_coords:
                    item = QTreeWidgetItem([
                        str(i),
                        name,
                        f"{lon:.6f}",
                        f"{lat:.6f}",
                        status
                    ])
                else:
                    item = QTreeWidgetItem([
                        str(i),
                        name,
                        "N/A",
                        "N/A",
                        status
                    ])
                self.tree.addTopLevelItem(item)
        except Exception as e:
            logger.error(f"更新地点表格失败: {str(e)}", exc_info=True)
            self.update_api_response(f"❌ 更新表格失败: {str(e)}")
    
    def search_scene_locations(self):
        """搜索指定城市、行政区和场景的地点"""
        try:
            # 获取选中的城市
            city = self.city_combo.currentData()
            if not city:
                QMessageBox.warning(self, "警告", "请选择城市")
                return
            
            # 获取选中的行政区
            selected_districts = []
            if self.all_districts_checkbox.isChecked():
                # 如果选择了全区域，不指定具体行政区
                selected_districts = [""]
            else:
                # 获取所有选中的行政区
                selected_districts = [district for district, checkbox in self.district_checkboxes.items() if checkbox.isChecked()]
                if not selected_districts:
                    QMessageBox.warning(self, "警告", "请至少选择一个行政区或选择全区域")
                    return
            
            # 获取选中的场景
            selected_scenes = [scene for scene, checkbox in self.scene_checkboxes.items() if checkbox.isChecked()]
            if not selected_scenes:
                QMessageBox.warning(self, "警告", "请至少选择一个生活场景")
                return
            
            # 显示搜索进度
            districts_str = "全区域" if self.all_districts_checkbox.isChecked() else ", ".join(selected_districts)
            self.status_label.setText(f"正在搜索{city}{districts_str}的{', '.join(selected_scenes)}...")
            self.status_label.setStyleSheet("color: blue")
            
            # 使用高德地图POI搜索API
            key = self.key_input.text().strip() if hasattr(self, 'key_input') else self.key
            
            # 存储搜索到的地点
            new_locations = []
            
            # 遍历选中的场景进行搜索
            for scene in selected_scenes:
                # 遍历选中的行政区进行搜索
                for district in selected_districts:
                    page = 1
                    total = 0
                    while True:
                        # 构建API请求URL
                        if district:
                            # 如果指定了行政区，将行政区添加到keywords中
                            search_keywords = f"{scene} {district}"
                        else:
                            search_keywords = scene
                        
                        url = f"https://restapi.amap.com/v3/place/text?keywords={quote(search_keywords)}&city={quote(city)}&output=json&offset=20&page={page}&key={key}"
                        
                        # 发送请求
                        response = requests.get(url)
                        data = response.json()
                        
                        # 检查API响应
                        if data.get('status') != '1':
                            QMessageBox.critical(self, "错误", f"搜索失败: {data.get('info')}")
                            return
                        
                        # 处理搜索结果
                        pois = data.get('pois', [])
                        if not pois:
                            break
                        
                        # 解析并添加地点
                        for poi in pois:
                            name = poi.get('name', '')
                            if name and name not in new_locations and name not in self.locations:
                                new_locations.append(name)
                        
                        # 获取总页数
                        if page == 1:
                            total = int(data.get('count', '0'))
                        
                        # 继续下一页或结束
                        page += 1
                        if page > (total // 20) + 1:
                            break
                        
                        # 避免API调用过于频繁
                        time.sleep(1)
            
            # 添加新地点到列表
            if new_locations:
                self.locations.extend(new_locations)
                self._update_location_table()
                self.status_label.setText(f"已添加{len(new_locations)}个地点，共{len(self.locations)}个地点")
                self.status_label.setStyleSheet("color: green")
                self.update_api_response(f"✅ 成功搜索到{len(new_locations)}个新地点")
            else:
                self.status_label.setText("未搜索到新地点")
                self.status_label.setStyleSheet("color: orange")
                self.update_api_response("ℹ️ 未搜索到新地点")
                
        except Exception as e:
            logger.error(f"搜索场景地点失败: {str(e)}", exc_info=True)
            self.update_api_response(f"❌ 搜索场景地点失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"搜索场景地点失败: {str(e)}")
    
    def on_city_changed(self):
        """城市选择变化时更新行政区选项"""
        # 清除所有已有的行政区多选框（除了全区域）
        for checkbox in self.district_checkboxes.values():
            checkbox.setParent(None)
        self.district_checkboxes.clear()
        
        # 获取选中的城市
        city = self.city_combo.currentData()
        
        if city and city in CITY_DISTRICTS:
            # 获取该城市的所有行政区
            districts = CITY_DISTRICTS[city]
            
            # 添加行政区多选框
            for district in districts:
                checkbox = QCheckBox(district)
                checkbox.stateChanged.connect(self.on_district_changed)
                self.district_layout.addWidget(checkbox)
                self.district_checkboxes[district] = checkbox
        
        # 重置全区域选项
        self.all_districts_checkbox.setChecked(True)
    
    def on_all_districts_changed(self, state):
        """全区域选项变化时处理其他选项的状态"""
        if state == Qt.Checked:
            # 如果全区域被选中，禁用所有其他行政区选项
            for checkbox in self.district_checkboxes.values():
                checkbox.setChecked(False)
                checkbox.setEnabled(False)
        else:
            # 如果全区域未被选中，启用所有其他行政区选项
            for checkbox in self.district_checkboxes.values():
                checkbox.setEnabled(True)
    
    def on_district_changed(self, state):
        """单个行政区选项变化时处理全区域选项的状态"""
        # 检查是否有任何行政区被选中
        any_district_checked = any(checkbox.isChecked() for checkbox in self.district_checkboxes.values())
        
        if any_district_checked:
            # 如果有任何行政区被选中，取消全区域选项
            self.all_districts_checkbox.setChecked(False)
    
    def delete_selected_locations(self):
        """删除选中的地点"""
        try:
            # 获取选中的项目
            selected_items = self.tree.selectedItems()
            if not selected_items:
                QMessageBox.warning(self, "警告", "请先选中要删除的地点")
                return
            
            # 确认删除
            reply = QMessageBox.question(self, "确认删除", f"确定要删除选中的{len(selected_items)}个地点吗？",
                                      QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            
            if reply == QMessageBox.Yes:
                # 收集要删除的地点名称
                locations_to_delete = [item.text(1) for item in selected_items]
                
                # 从locations列表中删除
                self.locations = [loc for loc in self.locations if loc not in locations_to_delete]
                
                # 从coordinates列表中删除
                self.coordinates = [coord for coord in self.coordinates if coord['name'] not in locations_to_delete]
                
                # 更新表格
                self._update_location_table()
                
                # 更新状态
                self.status_label.setText(f"已删除{len(locations_to_delete)}个地点，剩余{len(self.locations)}个地点")
                self.status_label.setStyleSheet("color: green")
                self.update_api_response(f"✅ 成功删除{len(locations_to_delete)}个地点")
        except Exception as e:
            logger.error(f"删除选中地点失败: {str(e)}", exc_info=True)
            self.update_api_response(f"❌ 删除选中地点失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"删除选中地点失败: {str(e)}")
    
    def show_all_locations_map(self):
        """显示所有地点及经纬度坐标在地图上"""
        try:
            # 获取所有有坐标的地点
            all_locations = []
            
            # 添加已获取坐标的地点
            for loc in self.coordinates:
                all_locations.append({
                    'name': loc['name'],
                    'lon': loc['lon'],
                    'lat': loc['lat'],
                    'status': "已获取坐标"
                })
            
            # 添加Excel导入的地点
            for loc in self.valid_locations:
                if not any(item['name'] == loc['name'] for item in all_locations):
                    all_locations.append({
                        'name': loc['name'],
                        'lon': loc['lon'],
                        'lat': loc['lat'],
                        'status': "Excel导入"
                    })
            
            if not all_locations:
                QMessageBox.warning(self, "警告", "没有可用的坐标数据，请先获取坐标")
                return
            
            self.update_api_response(f"🗺️ 开始生成所有地点地图，共 {len(all_locations)} 个地点")
            
            # 计算地图中心点
            all_lats = [loc['lat'] for loc in all_locations]
            all_lons = [loc['lon'] for loc in all_locations]
            center_lat = sum(all_lats) / len(all_lats)
            center_lon = sum(all_lons) / len(all_lons)
            
            # 创建高德地图
            map = folium.Map(
                location=[center_lat, center_lon],
                zoom_start=12,
                tiles='https://webrd03.is.autonavi.com/appmaptile?lang=zh_cn&size=1&scale=1&style=8&x={x}&y={y}&z={z}',
                attr='© <a href="https://ditu.amap.com/">高德地图</a>'
            )
            
            # 添加地点标记
            for i, loc in enumerate(all_locations):
                # 计算标记颜色
                if loc['status'] == "Excel导入":
                    color = "green"
                else:
                    color = "blue"
                
                # 创建标记
                folium.Marker(
                    location=[loc['lat'], loc['lon']],
                    tooltip=loc['name'],
                    popup=f"<b>{loc['name']}</b><br>经度: {loc['lon']:.6f}<br>纬度: {loc['lat']:.6f}<br>状态: {loc['status']}",
                    icon=folium.Icon(color=color, icon='info-sign', prefix='fa')
                ).add_to(map)
            
            # 保存地图文件
            temp_dir = tempfile.gettempdir()
            map_path = os.path.join(temp_dir, f"所有地点地图_{pd.Timestamp.now().strftime('%Y%m%d%H%M%S')}.html")
            map.save(map_path)
            
            # 自动打开地图
            webbrowser.open(f'file://{os.path.abspath(map_path)}')
            
            self.update_api_response(f"✅ 地图已生成并自动打开: {map_path}")
            self.status_label.setText(f"地图已生成，共显示 {len(all_locations)} 个地点")
            self.status_label.setStyleSheet("color: green")
            
        except Exception as e:
            logger.error(f"生成地图失败: {str(e)}", exc_info=True)
            self.update_api_response(f"❌ 生成地图失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"生成地图失败: {str(e)}")
    
    def import_locations_from_excel(self):
        """从Excel导入地点，支持直接导入经纬度坐标"""
        file_path, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls);;CSV文件 (*.csv)")
        
        if not file_path:
            return
        
        try:
            # 读取文件
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path, encoding='utf-8')
            else:
                df = pd.read_excel(file_path)
            
            self.update_api_response(f"📥 开始处理文件: {os.path.basename(file_path)}")
            self.update_api_response(f"📊 数据规模: {len(df)} 行, {len(df.columns)} 列")
            
            # 检测列名
            col_names = [col.lower().strip() for col in df.columns]
            self.update_api_response(f"📋 列名列表: {', '.join(col_names)}")
            
            # 检测经纬度列
            has_lat_lon = False
            lat_col_idx = None
            lon_col_idx = None
            name_col_idx = 0  # 默认第一列为名称列
            
            # 查找名称列
            for i, col in enumerate(col_names):
                if any(keyword in col for keyword in ['name', '地点', '地址', '名称']):
                    name_col_idx = i
                    break
            
            # 查找经度列
            for i, col in enumerate(col_names):
                if any(keyword in col for keyword in ['lon', '经度', 'lng', 'long']):
                    lon_col_idx = i
                    break
            
            # 查找纬度列
            for i, col in enumerate(col_names):
                if any(keyword in col for keyword in ['lat', '纬度']):
                    lat_col_idx = i
                    break
            
            # 检测是否有完整的经纬度信息
            if lon_col_idx is not None and lat_col_idx is not None:
                has_lat_lon = True
                self.update_api_response(f"✅ 检测到经纬度列: 经度列[{df.columns[lon_col_idx]}], 纬度列[{df.columns[lat_col_idx]}]")
                self.update_api_response(f"✅ 名称列: {df.columns[name_col_idx]}")
            
            # 清空现有数据
            imported_count = 0
            valid_coord_count = 0
            new_locations = []
            new_coordinates = []
            new_valid_locations = []
            
            # 处理每一行数据
            for idx, row in df.iterrows():
                name = str(row.iloc[name_col_idx]).strip()
                if not name or name == 'nan':
                    continue
                
                imported_count += 1
                location_data = {
                    'name': name,
                    'lon': None,
                    'lat': None
                }
                
                # 如果有经纬度信息，直接使用
                if has_lat_lon:
                    try:
                        lon = float(row.iloc[lon_col_idx])
                        lat = float(row.iloc[lat_col_idx])
                        
                        # 验证经纬度范围
                        if -180 <= lon <= 180 and -90 <= lat <= 90:
                            location_data['lon'] = lon
                            location_data['lat'] = lat
                            new_coordinates.append(location_data)
                            new_valid_locations.append(location_data)
                            valid_coord_count += 1
                            self.update_api_response(f"✅ 导入地点: {name} (直接使用Excel经纬度)")
                        else:
                            self.update_api_response(f"⚠️ 地点 {name} 经纬度超出范围，跳过")
                    except (ValueError, TypeError):
                        self.update_api_response(f"⚠️ 地点 {name} 经纬度格式错误，跳过")
                else:
                    new_locations.append(name)
                    self.update_api_response(f"✅ 导入地点: {name} (无经纬度)")
            
            # 更新全局数据
            self.locations.extend(new_locations)
            self.coordinates.extend(new_coordinates)
            self.valid_locations.extend(new_valid_locations)
            
            # 清空表格并添加新数据
            self.tree.clear()
            
            # 显示所有导入的地点，包括待查询的地点
            # 1. 先显示带有经纬度的地点
            lat_lon_items = []
            for loc in self.coordinates:
                lat_lon_items.append((loc['name'], True, loc['lon'], loc['lat']))
            
            # 2. 再显示待查询的地点
            for loc_name in self.locations:
                if not any(item[0] == loc_name for item in lat_lon_items):
                    lat_lon_items.append((loc_name, False, None, None))
            
            # 3. 按名称排序并显示
            for i, (name, has_coords, lon, lat) in enumerate(sorted(lat_lon_items, key=lambda x: x[0]), 1):
                if has_coords:
                    item = QTreeWidgetItem([
                        str(i),
                        name,
                        f"{lon:.6f}",
                        f"{lat:.6f}",
                        "✅ Excel导入"
                    ])
                else:
                    item = QTreeWidgetItem([
                        str(i),
                        name,
                        "N/A",
                        "N/A",
                        "⏳ 待查询"
                    ])
                self.tree.addTopLevelItem(item)
            
            # 更新状态
            status_msg = f"导入完成！共处理 {imported_count} 个地点"
            if has_lat_lon:
                status_msg += f"，其中 {valid_coord_count} 个包含有效经纬度"
            
            self.status_label.setText(status_msg)
            self.status_label.setStyleSheet("color: green")
            self.update_api_response(f"{status_msg}")
            self.update_api_response(f"📊 当前有效地点: {len(self.valid_locations)}")
            self.update_api_response(f"📊 当前待查询地点: {len(self.locations)}")
            
        except Exception as e:
            self.update_api_response(f"❌ 导入错误: {str(e)}")
            QMessageBox.warning(self, "导入错误", f"导入失败: {str(e)}")
    
    def fetch_coordinates(self):
        """批量获取坐标"""
        if not self.locations:
            QMessageBox.showwarning("警告", "请先添加地点")
            return
        
        # 确保UI控件存在
        if not hasattr(self, 'fetch_btn') or not hasattr(self, 'status_label') or not hasattr(self, 'city_combo'):
            QMessageBox.critical(self, "错误", "UI控件初始化失败")
            return
        
        # 在主线程中获取城市名，避免在线程中访问UI控件
        city = self.city_combo.currentText().strip() or '全国'
        
        self.fetch_btn.setEnabled(False)
        # 将批量获取坐标状态显示在处理日志下方
        if hasattr(self, 'fetch_status_label'):
            self.fetch_status_label.setText("正在批量获取坐标...")
            self.fetch_status_label.setStyleSheet("color: blue")
        # 同时清理右侧状态栏，避免重复提示
        if hasattr(self, 'status_label'):
            self.status_label.setText("")
        self.update_api_response(f"开始获取 {len(self.locations)} 个地点的坐标...")
        
        try:
            # 启动线程获取坐标，并传递城市名参数
            threading.Thread(target=self._fetch_coordinates_thread, args=(city,), daemon=True).start()
        except Exception as e:
            self.status_label.setText(f"启动线程失败: {str(e)}")
            self.status_label.setStyleSheet("color: red")
            self.fetch_btn.setEnabled(True)
            logger.error(f"启动坐标获取线程失败: {str(e)}", exc_info=True)
    
    def _fetch_coordinates_thread(self, city):
        """在线程中获取坐标"""
        try:
            # 线程安全：保存当前需要处理的地点列表
            current_locations = self.locations.copy()
            
            base_url = "https://restapi.amap.com/v3/geocode/geo"
            
            # 用于存储新获取的坐标
            new_coords = {}
            
            for i, location in enumerate(current_locations, 1):
                try:
                    params = {
                        'address': location,
                        'key': self.key,
                        'city': city
                    }
                    
                    response = requests.get(base_url, params=params, timeout=10)
                    response.raise_for_status()
                    data = response.json()
                    
                    if data['status'] == '1' and data['geocodes']:
                        geocode = data['geocodes'][0]
                        lon, lat = map(float, geocode['location'].split(','))
                        new_coords[location] = {'lon': lon, 'lat': lat, 'status': "✅ 成功"}
                        
                        self.update_api_response(f"✅ [{i}/{len(current_locations)}] {location} - 获取成功")
                    else:
                        new_coords[location] = {'lon': None, 'lat': None, 'status': "❌ 失败"}
                        self.update_api_response(f"❌ [{i}/{len(current_locations)}] {location} - 未找到")
                    
                    # 适当延迟，避免API调用过于频繁
                    time.sleep(0.5)
                    
                except requests.exceptions.RequestException as e:
                    new_coords[location] = {'lon': None, 'lat': None, 'status': "❌ 网络错误"}
                    self.update_api_response(f"❌ [{i}/{len(current_locations)}] {location} - 网络错误: {str(e)}")
                    time.sleep(1)
                except ValueError as e:
                    new_coords[location] = {'lon': None, 'lat': None, 'status': "❌ 解析错误"}
                    self.update_api_response(f"❌ [{i}/{len(current_locations)}] {location} - 解析错误: {str(e)}")
                except Exception as e:
                    new_coords[location] = {'lon': None, 'lat': None, 'status': "❌ 错误"}
                    self.update_api_response(f"❌ [{i}/{len(current_locations)}] {location} - 错误: {str(e)}")
            
            # 所有坐标获取完成后，在主线程中更新UI和数据
            self._update_coordinates_ui(new_coords)
            
        except Exception as e:
            # 捕获所有异常，确保线程不会崩溃
            error_msg = f"获取坐标失败: {str(e)}"
            logger.error(error_msg, exc_info=True)
            self.update_api_response(f"❌ 严重错误: {str(e)}")
        finally:
            # 确保按钮状态在主线程中恢复
            if hasattr(self, 'fetch_btn'):
                self.fetch_btn.setEnabled(True)
            if hasattr(self, 'status_label'):
                self.status_label.setStyleSheet("color: black")
    
    def _update_coordinates_ui(self, new_coords):
        """在主线程中更新坐标UI和数据"""
        try:
            # 更新坐标数据
            # 1. 先保存现有的坐标数据
            existing_coords = {loc['name']: loc for loc in self.coordinates}
            
            # 2. 更新或添加新获取的坐标
            updated_coords = []
            updated_valid_locations = []
            success_count = 0
            
            for location, coord_info in new_coords.items():
                if coord_info['lon'] is not None and coord_info['lat'] is not None:
                    loc_data = {
                        'name': location,
                        'lon': coord_info['lon'],
                        'lat': coord_info['lat']
                    }
                    updated_coords.append(loc_data)
                    updated_valid_locations.append(loc_data)
                    success_count += 1
                    existing_coords[location] = loc_data
                elif location in existing_coords:
                    updated_coords.append(existing_coords[location])
                    updated_valid_locations.append(existing_coords[location])
            
            # 3. 添加其他已存在但不在本次更新中的坐标
            for loc in self.coordinates:
                if loc['name'] not in new_coords:
                    updated_coords.append(loc)
                    updated_valid_locations.append(loc)
            
            # 更新全局数据
            self.coordinates = updated_coords
            self.valid_locations = updated_valid_locations
            
            # 更新表格UI
            # 1. 清空表格
            self.tree.clear()
            
            # 2. 显示所有地点，包括已获取坐标和待查询的
            all_items = []
            
            # 添加已获取坐标的地点
            for loc in updated_coords:
                all_items.append((loc['name'], True, loc['lon'], loc['lat'], "✅ 获取成功"))
            
            # 添加待查询的地点
            for location in self.locations:
                if not any(item[0] == location for item in all_items):
                    all_items.append((location, False, None, None, "⏳ 待查询"))
            
            # 3. 添加Excel导入的地点
            for loc in self.coordinates:
                if not any(item[0] == loc['name'] for item in all_items):
                    all_items.append((loc['name'], True, loc['lon'], loc['lat'], "✅ Excel导入"))
            
            # 4. 去重并按名称排序
            seen_names = set()
            unique_items = []
            for item in all_items:
                if item[0] not in seen_names:
                    unique_items.append(item)
                    seen_names.add(item[0])
            
            # 5. 显示到表格
            for i, (name, has_coords, lon, lat, status) in enumerate(sorted(unique_items, key=lambda x: x[0]), 1):
                if has_coords:
                    item = QTreeWidgetItem([
                        str(i),
                        name,
                        f"{lon:.6f}",
                        f"{lat:.6f}",
                        status
                    ])
                else:
                    item = QTreeWidgetItem([
                        str(i),
                        name,
                        "N/A",
                        "N/A",
                        status
                    ])
                self.tree.addTopLevelItem(item)
            
            # 更新状态（显示在处理日志下方）
            total = len(new_coords)
            if hasattr(self, 'fetch_status_label'):
                self.fetch_status_label.setText(f"完成！成功获取 {success_count}/{total} 个地点的坐标 (成功率: {success_count/total:.1%})")
                self.fetch_status_label.setStyleSheet("color: green")
            else:
                self.status_label.setText(f"完成！成功获取 {success_count}/{total} 个地点的坐标 (成功率: {success_count/total:.1%})")
                self.status_label.setStyleSheet("color: green")
            self.update_api_response(f"获取完成！成功率: {success_count}/{total}")
            # 根据有效坐标数量控制生成路线按钮的可用性（至少需要2个点）
            try:
                if hasattr(self, 'generate_route_btn'):
                    self.generate_route_btn.setEnabled(len(self.valid_locations) >= 2)
            except Exception:
                pass
            
        except Exception as e:
            error_msg = f"更新坐标UI失败: {str(e)}"
            logger.error(error_msg, exc_info=True)
            self.status_label.setText("更新坐标UI失败")
            self.status_label.setStyleSheet("color: red")
            self.update_api_response(f"❌ 更新UI错误: {str(e)}")
        finally:
            # 确保按钮状态恢复
            if hasattr(self, 'fetch_btn'):
                self.fetch_btn.setEnabled(True)
    
    def update_api_response(self, message):
        """更新日志显示（线程安全）"""
        from PyQt5.QtCore import QMetaObject, Qt, Q_ARG
        
        def safe_update():
            """在主线程中安全更新UI"""
            try:
                self.response_text.append(message)
                self.response_text.verticalScrollBar().setValue(self.response_text.verticalScrollBar().maximum())
            except Exception as e:
                logger.error(f"更新日志失败: {str(e)}", exc_info=True)
        
        # 使用QMetaObject.invokeMethod确保在主线程中执行UI更新
        QMetaObject.invokeMethod(
            self.response_text,
            "append",
            Qt.QueuedConnection,
            Q_ARG(str, message)
        )
        
        # 确保垂直滚动条滚动到底部
        QMetaObject.invokeMethod(
            self.response_text.verticalScrollBar(),
            "setValue",
            Qt.QueuedConnection,
            Q_ARG(int, self.response_text.verticalScrollBar().maximum())
        )
    
    def calculate_distance_between_points(self, point1, point2):
        """计算两个坐标点之间的距离（单位：公里）"""
        EARTH_RADIUS = 6371
        
        lat1_rad = math.radians(point1['lat'])
        lon1_rad = math.radians(point1['lon'])
        lat2_rad = math.radians(point2['lat'])
        lon2_rad = math.radians(point2['lon'])
        
        dlat = lat2_rad - lat1_rad
        dlon = lon2_rad - lon1_rad
        
        a = math.sin(dlat / 2) ** 2 + math.cos(lat1_rad) * math.cos(lat2_rad) * math.sin(dlon / 2) ** 2
        c = 2 * math.asin(math.sqrt(a))
        distance = EARTH_RADIUS * c
        
        return distance
    
    def is_waypoint_in_valid_range(self, waypoint, start_point, end_point, other_waypoints=None):
        """检查途径点是否在合理的距离范围内"""
        config = self.route_config
        
        dist_to_start = self.calculate_distance_between_points(waypoint, start_point)
        if not (config['waypoint_min_distance'] <= dist_to_start <= config['waypoint_max_distance']):
            return False
        
        dist_to_end = self.calculate_distance_between_points(waypoint, end_point)
        if not (config['waypoint_min_distance'] <= dist_to_end <= config['waypoint_max_distance']):
            return False
        
        if other_waypoints:
            for other_wp in other_waypoints:
                dist_between = self.calculate_distance_between_points(waypoint, other_wp)
                if not (config['between_waypoint_min'] <= dist_between <= config['between_waypoint_max']):
                    return False
        
        if len(other_waypoints or []) > 0:
            for other_wp in (other_waypoints or []):
                if self.are_points_collinear(start_point, waypoint, other_wp):
                    return False
        
        return True
    
    def are_points_collinear(self, p1, p2, p3, tolerance=0.05):
        """检查三个点是否共线"""
        area = abs(
            (p2['lat'] - p1['lat']) * (p3['lon'] - p1['lon']) -
            (p3['lat'] - p1['lat']) * (p2['lon'] - p1['lon'])
        )
        
        return area < tolerance
    
    def calculate_route_signature(self, route):
        """生成路线的指纹（特征值）"""
        all_points = [
            (route['start_point']['lon'], route['start_point']['lat']),
            (route['end_point']['lon'], route['end_point']['lat'])
        ]
        for wp in route.get('waypoint_details', []):
            all_points.append((wp['lon'], wp['lat']))
        
        total_distance = 0
        for i in range(len(all_points) - 1):
            p1 = {'lat': all_points[i][1], 'lon': all_points[i][0]}
            p2 = {'lat': all_points[i+1][1], 'lon': all_points[i+1][0]}
            total_distance += self.calculate_distance_between_points(p1, p2)
        
        center_lat = sum(p[1] for p in all_points) / len(all_points)
        center_lon = sum(p[0] for p in all_points) / len(all_points)
        
        signature = {
            'total_distance': round(total_distance, 2),
            'center': (round(center_lat, 4), round(center_lon, 4)),
            'point_count': len(all_points),
            'points_sorted': tuple(sorted(all_points))
        }
        
        return signature
    
    def calculate_route_similarity(self, route1, route2):
        """计算两条路线的相似度"""
        sig1 = self.calculate_route_signature(route1)
        sig2 = self.calculate_route_signature(route2)
        
        points1 = set(sig1['points_sorted'])
        points2 = set(sig2['points_sorted'])
        overlap = len(points1 & points2)
        total = len(points1 | points2)
        point_similarity = overlap / total if total > 0 else 0
        
        dist_diff = abs(sig1['total_distance'] - sig2['total_distance'])
        avg_dist = (sig1['total_distance'] + sig2['total_distance']) / 2
        distance_similarity = 1 - (dist_diff / avg_dist if avg_dist > 0 else 0)
        
        center1 = {'lat': sig1['center'][0], 'lon': sig1['center'][1]}
        center2 = {'lat': sig2['center'][0], 'lon': sig2['center'][1]}
        center_dist = self.calculate_distance_between_points(center1, center2)
        center_similarity = max(0, 1 - (center_dist / 5))
        
        final_similarity = (
            point_similarity * 0.5 +
            distance_similarity * 0.3 +
            center_similarity * 0.2
        )
        
        return final_similarity
    
    def is_route_duplicate(self, new_route, existing_routes):
        """检查新路线是否与现有路线重复"""
        if not self.route_config['enable_deduplication']:
            return False
        
        if not existing_routes:
            return False
        
        for existing_route in existing_routes:
            similarity = self.calculate_route_similarity(new_route, existing_route)
            if similarity > self.route_config['similarity_threshold']:
                self.update_api_response(
                    f"⚠️ 新路线与已生成的路线 {existing_route['route_id']} 相似度过高 "
                    f"({similarity:.2%})，已过滤"
                )
                return True
        
        return False
    
    def select_optimal_waypoints(self, start_point, end_point, waypoint_num, used_waypoints_set):
        """智能选择最优的途径点"""
        config = self.route_config
        
        self.update_api_response(f"🔍 开始选择 {waypoint_num} 个途径点")
        self.update_api_response(f"📋 候选点总数: {len(self.valid_locations)}")
        
        candidates = []
        for point in self.valid_locations:
            if point['name'] == start_point['name'] or point['name'] == end_point['name']:
                continue
            if point['name'] in used_waypoints_set:
                continue
            # 必须有经纬度
            if point.get('lat') is None or point.get('lon') is None:
                continue
            candidates.append(point)
        
        self.update_api_response(f"🔢 可用候选点: {len(candidates)}")
        
        # 如果候选点不足，尝试从所有有效地点补充最近的点（按距离起点）
        if len(candidates) < waypoint_num:
            # 搜索可补充的候选（排除已用、起终点和已有candidates）
            supplement_pool = [p for p in self.valid_locations
                               if p['name'] not in used_waypoints_set
                               and p['name'] != start_point['name']
                               and p['name'] != end_point['name']
                               and p not in candidates
                               and p.get('lat') is not None and p.get('lon') is not None]
            # 按距起点距离排序并补充
            supplement_pool.sort(key=lambda p: self.calculate_distance_between_points(p, start_point))
            need = waypoint_num - len(candidates)
            for p in supplement_pool[:need]:
                candidates.append(p)

            self.update_api_response(
                f"⚠️ 候选途径点不足，已自动补充至 {len(candidates)} 个（目标 {waypoint_num}）。"
            )
            # 若仍不足，返回现有的全部
            if len(candidates) <= 0:
                return []
        
        # 使用更严格的相邻距离约束：200m - 1000m（0.2km - 1.0km）
        min_adj_km = 0.5
        max_adj_km = 1.5

        valid_candidates = []
        for candidate in candidates:
            # 检查与起点和终点距离在全局路由配置范围内
            if not self.is_waypoint_in_valid_range(candidate, start_point, end_point):
                continue
            # 作为单独候选，先接受
            valid_candidates.append(candidate)
        
        self.update_api_response(f"✅ 符合距离条件的候选点: {len(valid_candidates)}")
        
        if len(valid_candidates) < waypoint_num:
            self.update_api_response(
                f"⚠️ 警告：符合距离条件的途径点不足 ({len(valid_candidates)}/{waypoint_num})"
            )
            return valid_candidates[:waypoint_num]
        
        selected_waypoints = []
        remaining_candidates = valid_candidates.copy()

        # 选择逻辑：按与起点的距离升序选择，保证顺序从起点到终点（避免回头），并满足相邻距离范围
        # 先按距离起点排序所有候选
        remaining_candidates.sort(key=lambda p: self.calculate_distance_between_points(p, start_point))

        for candidate in remaining_candidates:
            if len(selected_waypoints) >= waypoint_num:
                break
            # 如果没有前一个点，则只需检查与起点距离约束（已在 is_waypoint_in_valid_range 中检查）
            if not selected_waypoints:
                selected_waypoints.append(candidate)
                continue

            prev = selected_waypoints[-1]
            adj_dist = self.calculate_distance_between_points(prev, candidate)
            # 检查相邻点距离在 0.2km - 1.0km 范围内
            if adj_dist < min_adj_km or adj_dist > max_adj_km:
                # 距离不合适，跳过
                continue
            # 额外检查与start/end的全局合理性
            if not self.is_waypoint_in_valid_range(candidate, start_point, end_point, [prev]):
                continue
            selected_waypoints.append(candidate)

        # 若选出的数量依然不足，尝试从剩余候选中补齐最近点（不再严格要求邻距），以满足数量
        if len(selected_waypoints) < waypoint_num:
            need = waypoint_num - len(selected_waypoints)
            extras = [p for p in remaining_candidates if p not in selected_waypoints]
            extras.sort(key=lambda p: self.calculate_distance_between_points(p, start_point))
            for p in extras[:need]:
                selected_waypoints.append(p)
        
        self.update_api_response(f"📌 最终选择的途径点: {len(selected_waypoints)}")
        return selected_waypoints
    
    def generate_navigation_url(self, start_point, end_point, waypoints):
        """生成高德导航链接"""
        try:
            base_url = "https://ditu.amap.com/dir?type=car&policy=1"
            
            start_lnglat = f"{start_point['lon']},{start_point['lat']}"
            base_url += f"&from[lnglat]={start_lnglat}"
            base_url += f"&from[name]={quote(start_point['name'])}"
            
            end_lnglat = f"{end_point['lon']},{end_point['lat']}"
            base_url += f"&to[lnglat]={end_lnglat}"
            base_url += f"&to[name]={quote(end_point['name'])}"
            
            for i, point in enumerate(waypoints):
                wp_lnglat = f"{point['lon']},{point['lat']}"
                base_url += f"&via[{i}][lnglat]={wp_lnglat}"
                base_url += f"&via[{i}][name]={quote(point['name'])}"
                
            return base_url
        except Exception as e:
            self.update_api_response(f"❌ 生成导航链接错误: {str(e)}")
            logger.error(f"生成导航链接错误: {str(e)}")
            return None
    
    def get_driving_route(self, start, end, waypoints):
        """获取驾驶路线"""
        try:
            origin = f"{start['lon']},{start['lat']}"
            destination = f"{end['lon']},{end['lat']}"
            
            waypoint_str = ""
            if waypoints:
                waypoint_str = ";".join([f"{wp['lon']},{wp['lat']}" for wp in waypoints])
            
            route_url = "https://restapi.amap.com/v3/direction/driving"
            params = {
                'origin': origin,
                'destination': destination,
                'waypoints': waypoint_str if waypoint_str else None,
                'key': self.key
            }
            
            response = requests.get(route_url, params=params, timeout=10)
            data = response.json()
            
            if data['status'] == '1' and 'route' in data:
                route = data['route']
                points = []
                
                for path in route['paths']:
                    for step in path['steps']:
                        for point in step['polyline'].split(';'):
                            if point:
                                lon, lat = point.split(',')
                                points.append([float(lat), float(lon)])
                
                return points if points else None
            
        except Exception as e:
            logger.error(f"获取驾驶路线错误: {str(e)}")
        
        return None
    
    def generate_simple_route(self, start_point, end_point, waypoints):
        """当无法从API获取路线时，生成简单的直线路径"""
        points = [[start_point['lat'], start_point['lon']]]
        
        for wp in waypoints:
            points.append([wp['lat'], wp['lon']])
        
        points.append([end_point['lat'], end_point['lon']])
        
        return points
    
    def generate_route(self, route_num, waypoint_num, existing_routes=None):
        """【改进版】生成一条测试路线"""
        if existing_routes is None:
            existing_routes = []
        
        if len(self.valid_locations) < 2:
            self.update_api_response("❌ 错误: 有效地点数量不足，无法生成路线")
            return None
        
        max_attempts = 100
        for attempt in range(max_attempts):
            start_idx = random.randint(0, len(self.valid_locations) - 1)
            end_idx = random.randint(0, len(self.valid_locations) - 1)
            
            if start_idx != end_idx:
                break
        else:
            self.update_api_response("❌ 错误: 无法找到不同的起点和终点")
            return None
        
        start_point = self.valid_locations[start_idx]
        end_point = self.valid_locations[end_idx]
        
        straight_distance = self.calculate_distance_between_points(start_point, end_point)
        self.update_api_response(
            f"📍 路线 {route_num}: 起点[{start_point['name']}] → 终点[{end_point['name']}] "
            f"(直线距离: {straight_distance:.2f}km)"
        )
        
        used_waypoints_set = set()
        for existing_route in existing_routes:
            used_waypoints_set.add(existing_route['start_point']['name'])
            used_waypoints_set.add(existing_route['end_point']['name'])
            for wp_name in existing_route['waypoints'].split('; '):
                if wp_name.strip():
                    used_waypoints_set.add(wp_name.strip())
        
        waypoints = []
        if waypoint_num > 0:
            waypoints = self.select_optimal_waypoints(
                start_point, end_point, waypoint_num, used_waypoints_set
            )
            
            if waypoints:
                waypoint_dists = []
                for wp in waypoints:
                    dist = self.calculate_distance_between_points(start_point, wp)
                    waypoint_dists.append(f"{wp['name']}({dist:.2f}km)")
                
                self.update_api_response(
                    f"   └─ 途径点: {' → '.join(waypoint_dists)}"
                )
        
        nav_url = self.generate_navigation_url(start_point, end_point, waypoints)
        if not nav_url:
            self.update_api_response(f"❌ 路线 {route_num} 生成导航链接失败")
            return None
        
        real_points = self.get_driving_route(start_point, end_point, waypoints)
        
        if not real_points:
            waypoint_coords = [{"lat": wp["lat"], "lon": wp["lon"]} for wp in waypoints]
            real_points = self.generate_simple_route(start_point, end_point, waypoint_coords)
        
        route_info = {
            'route_id': route_num,
            'start_point': start_point,
            'end_point': end_point,
            'waypoints': '; '.join([wp['name'] for wp in waypoints]),
            'navigation_url': nav_url,
            'real_points': real_points if real_points else [],
            'waypoint_details': [{'name': wp['name'], 'lat': wp['lat'], 'lon': wp['lon']} for wp in waypoints],
            'straight_distance': straight_distance,
            'waypoint_count': len(waypoints)
        }
        
        if self.is_route_duplicate(route_info, existing_routes):
            self.update_api_response(f"⏭️ 路线 {route_num} 已被过滤（与现有路线过于相似）")
            return None
        
        self.update_api_response(f"✅ 路线 {route_num} 已成功生成")
        return route_info
    
    def generate_routes(self):
        """【改进版】批量生成路线（支持去重和智能选择）"""
        try:
            waypoint_num = self.waypoint_spin.value()
            target_route_num = self.route_num_spin.value()
            
            self.update_api_response(f"\n{'='*50}")
            self.update_api_response(f"开始生成 {target_route_num} 条测试路线")
            self.update_api_response(f"每条路线包含 {waypoint_num} 个途径点")
            self.update_api_response(f"路线去重: {'已启用' if self.route_config['enable_deduplication'] else '已禁用'}")
            self.update_api_response(f"相似度阈值: {self.route_config['similarity_threshold']:.2%}")
            self.update_api_response(f"途径点距离范围: {self.route_config['waypoint_min_distance']}-")
            self.update_api_response(f"{self.route_config['waypoint_max_distance']}km")
            self.update_api_response(f"{'='*50}\n")
            
            self.route_data = []
            failed_count = 0
            max_failed = 10
            
            route_id = 1
            while len(self.route_data) < target_route_num and failed_count < max_failed:
                route = self.generate_route(route_id, waypoint_num, self.route_data)
                
                if route:
                    self.route_data.append(route)
                    self.status_label.setText(
                        f"已生成 {len(self.route_data)}/{target_route_num} 条有效路线"
                    )
                    failed_count = 0
                    time.sleep(1)
                else:
                    failed_count += 1
                
                route_id += 1
            
            if self.generate_realistic_route_map():
                self.update_api_response("✅ 综合路线地图生成成功")
                self.all_routes_button.setEnabled(True)
            
            self.status_label.setText(
                f"✅ 完成! 成功生成 {len(self.route_data)} 条路线"
            )
            self.status_label.setStyleSheet("color: green")
            self.update_api_response(f"\n{'='*50}")
            self.update_api_response(f"生成完成! 共生成 {len(self.route_data)} 条有效路线")
            self.update_api_response(f"{'='*50}\n")
            
        except Exception as e:
            self.status_label.setText(f"❌ 路线生成错误: {str(e)}")
            self.status_label.setStyleSheet("color: red")
            self.update_api_response(f"❌ 路线生成错误: {str(e)}")
            logger.error(f"路线生成错误: {str(e)}")
        finally:
            self.generate_route_btn.setEnabled(True)
            self.fetch_btn.setEnabled(True)
            self.import_btn.setEnabled(True)
    
    def generate_realistic_route_map(self):
        """生成综合路线地图"""
        try:
            if not self.route_data:
                return False
            
            all_lats = []
            all_lons = []
            
            for route in self.route_data:
                for point in route.get('real_points', []):
                    all_lats.append(point[0])
                    all_lons.append(point[1])
            
            if not all_lats:
                return False
            
            center_lat = sum(all_lats) / len(all_lats)
            center_lon = sum(all_lons) / len(all_lons)
            
            route_map = folium.Map(
                location=[center_lat, center_lon],
                zoom_start=12,
                tiles='https://webrd03.is.autonavi.com/appmaptile?lang=zh_cn&size=1&scale=1&style=8&x={x}&y={y}&z={z}',
                attr='© <a href="https://ditu.amap.com/">高德地图</a>'
            )
            
            colors = ['red', 'blue', 'green', 'purple', 'orange', 'darkred', 'darkblue', 'darkgreen']
            
            for i, route in enumerate(self.route_data):
                color = colors[i % len(colors)]
                
                if route.get('real_points'):
                    folium.PolyLine(
                        route['real_points'],
                        color=color,
                        weight=3,
                        opacity=0.8,
                        popup=f"路线 {route['route_id']}"
                    ).add_to(route_map)
            
            temp_dir = tempfile.gettempdir()
            self.combined_map_path = os.path.join(temp_dir, 'all_routes_map.html')
            route_map.save(self.combined_map_path)
            
            return True
            
        except Exception as e:
            logger.error(f"生成综合地图错误: {str(e)}")
            return False
    
    def view_all_routes_map(self):
        """查看所有路线地图"""
        if self.combined_map_path and os.path.exists(self.combined_map_path):
            webbrowser.open('file://' + os.path.realpath(self.combined_map_path))
        else:
            QMessageBox.information(self, "提示", "请先生成路线")
    
    def start_generating_routes(self):
        """【关键】启动路线生成的主方法"""
        if not self.valid_locations or len(self.valid_locations) < 2:
            QMessageBox.showwarning("警告", "有效地点数量不足，无法生成路线")
            return
        
        waypoint_num = self.waypoint_spin.value()
        route_num = self.route_num_spin.value()
        
        required_points = route_num * (waypoint_num + 2)
        if required_points > len(self.valid_locations):
            reply = QMessageBox.question(self,
                                       "确认", 
                                       f"生成 {route_num} 条路线需要至少 {required_points} 个有效地点，但只有 {len(self.valid_locations)} 个。\n"
                                       "可能无法生成所有路线，是否继续？",
                                       QMessageBox.Yes | QMessageBox.No,
                                       QMessageBox.No)
            if reply != QMessageBox.Yes:
                return
        
        self.generate_route_btn.setEnabled(False)
        self.fetch_btn.setEnabled(False)
        self.import_btn.setEnabled(False)
        
        threading.Thread(target=self.generate_routes, daemon=True).start()
    
    def export_excel(self):
        """导出结果到Excel"""
        if not self.locations and not self.route_data:
            QMessageBox.showwarning("警告", "没有可导出的数据")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(self, "保存Excel文件", "", "Excel文件 (*.xlsx)")
        
        if not file_path:
            return
        
        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # Sheet 1: 坐标信息
                if self.coordinates:
                    coord_df = pd.DataFrame(self.coordinates)
                    coord_df.to_excel(writer, sheet_name="坐标信息", index=False)
                    self.update_api_response(f"✅ 已导出坐标结果: {len(self.coordinates)}条记录")
                
                # Sheet 2: 路线信息
                if self.route_data:
                    route_list = []
                    for route in self.route_data:
                        route_info = {
                            '路线ID': route['route_id'],
                            '起点': route['start_point']['name'],
                            '起点经度': route['start_point']['lon'],
                            '起点纬度': route['start_point']['lat'],
                            '终点': route['end_point']['name'],
                            '终点经度': route['end_point']['lon'],
                            '终点纬度': route['end_point']['lat'],
                            '途径点': route['waypoints'],
                            '途径点数量': route['waypoint_count'],
                            '直线距离(km)': route['straight_distance'],
                            '导航链接': route['navigation_url']
                        }
                        route_list.append(route_info)
                    
                    route_df = pd.DataFrame(route_list)
                    route_df.to_excel(writer, sheet_name="路线信息", index=False)
                    self.update_api_response(f"✅ 已导出路线结果: {len(route_list)}条记录")
            
            self.update_api_response(f"✅ 所有数据已成功导出到: {file_path}")
            QMessageBox.information(self, "导出成功", f"数据已成功导出到: {file_path}")
        except Exception as e:
            self.update_api_response(f"❌ 导出失败: {str(e)}")
            QMessageBox.critical(self, "导出失败", f"导出失败: {str(e)}")
    
    def import_json(self):
        """批量导入JSON文件，支持增量添加"""
        file_paths, _ = QFileDialog.getOpenFileNames(self, "选择JSON文件", "", "JSON文件 (*.json)")
        if file_paths:
            added_count = 0
            for file_path in file_paths:
                if file_path in self.json_files:
                    continue  # 已经添加过，跳过
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    # 检查格式
                    if isinstance(data, list) and len(data) > 0 and 'pointList' in data[0]:
                        route_data = data[0]
                        route_name = route_data.get('routeName', os.path.basename(file_path))
                        self.json_files.append(file_path)
                        self.json_data_list.append(route_data)
                        self.route_list.addItem(f"{route_name} ({os.path.basename(file_path)})")
                        added_count += 1
                    else:
                        raise ValueError(f"{file_path} 格式不正确")
                except Exception as e:
                    self.log_text.append(f"{file_path} 错误: {str(e)}")
            if self.json_files:
                self.import_status.setText(f"已导入 {len(self.json_files)} 个JSON文件，本次新增 {added_count} 个")
                self.calculate_btn.setEnabled(True)
            else:
                self.import_status.setText("未成功导入任何JSON文件")
                self.calculate_btn.setEnabled(False)
    
    def calculate_route(self):
        """批量计算所有JSON路线"""
        if not self.json_data_list:
            QMessageBox.warning(self, "警告", "请先导入JSON文件")
            return
        self.routes_result = []
        self.progress_bar.setValue(0)
        self.calculate_btn.setEnabled(False)
        self.import_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        self.log_text.append("开始批量计算路线...")
        self._calculate_next_json(0)

    def _calculate_next_json(self, idx):
        if idx >= len(self.json_data_list):
            self.progress_bar.setValue(100)
            self.import_status.setText("所有路线计算完成")
            self.calculate_btn.setEnabled(True)
            self.import_btn.setEnabled(True)
            self.export_btn.setEnabled(True)
            self.log_text.append("所有路线计算完成")
            return
        route_data = self.json_data_list[idx]
        waypoints = [f"{p['lon']},{p['lat']}" for p in route_data['pointList']]
        key = self.key_input.text().strip()
        self.import_status.setText(f"正在计算第{idx+1}/{len(self.json_data_list)}个路线...")
        self.calc_thread = RouteCalculator(waypoints, key, self.backup_keys)
        self.calc_thread.progress_updated.connect(self.update_progress)
        self.calc_thread.calculation_finished.connect(lambda all_coords, segs, road_types, road_names: self._on_single_calc_finished(idx, all_coords, segs, road_types, road_names))
        self.calc_thread.error_occurred.connect(self.on_calculation_error)
        self.calc_thread.start()

    def _on_single_calc_finished(self, idx, all_coords, segs, road_types, road_names):
        # 调试：打印道路类型信息
        print(f"路线 {self.json_files[idx]} 计算完成，获取到 {len(road_types)} 个道路类型信息")
        type_counts = {}
        for rt in road_types:
            type_counts[rt] = type_counts.get(rt, 0) + 1
        for rt, count in type_counts.items():
            print(f"  类型 {rt}: {count}个点")
        
        self.routes_result.append({
            "json_file": self.json_files[idx],
            "all_coordinates": all_coords,
            "route_segments": segs,
            "route_name": self.json_data_list[idx].get("routeName", os.path.basename(self.json_files[idx])),
            "road_types": road_types,
            "road_names": road_names
        })
        self.log_text.append(f"{self.json_files[idx]} 计算完成")
        self._calculate_next_json(idx + 1)
        
        # 检查是否所有计算都完成了
        if len(self.routes_result) == len(self.json_files):
            self.export_btn.setEnabled(True)
            self.log_text.append("所有路线计算完成")
            self.progress_bar.setValue(0)
    
    def update_progress(self, value):
        """更新进度条"""
        self.progress_bar.setValue(value)
    
    def on_calculation_error(self, error_msg):
        """计算错误的回调"""
        QMessageBox.critical(self, "计算错误", f"计算路线时出错: {error_msg}")
        self.import_status.setText("计算出错")
        self.calculate_btn.setEnabled(True)
        self.import_btn.setEnabled(True)
        self.log_text.append(f"错误: {error_msg}")
    
    def export_all(self):
        """批量导出所有Excel文件"""
        #if not hasattr(self, "routes_result") or not self.routes_result:
            #QMessageBox.warning(self, "警告", "没有可保存的路线数据")
            #return
        
        # 检查是否自动导出
        if hasattr(self, 'auto_export_dir') and self.auto_export_dir:
            output_dir = self.auto_export_dir
            # 清除自动导出标志，避免影响下次手动导出
            del self.auto_export_dir
        else:
            output_dir = QFileDialog.getExistingDirectory(self, "选择输出目录")
            if not output_dir:
                return
        for result in self.routes_result:
            base_name = os.path.splitext(os.path.basename(result["json_file"]))[0]
            excel_path = os.path.join(output_dir, f"{base_name}.xlsx")
            try:
                with pd.ExcelWriter(excel_path) as writer:
                    # 导出所有坐标点
                    coords_data = []
                    
                    # 调试：打印道路类型信息
                    if "road_types" in result:
                        road_type_counts = {}
                        for rt in result["road_types"]:
                            road_type_counts[rt] = road_type_counts.get(rt, 0) + 1
                        print(f"导出文件 {base_name} 的道路类型统计:")
                        for rt, count in road_type_counts.items():
                            print(f"类型 '{rt}': {count}个点")
                    else:
                        print(f"导出文件 {base_name} 没有道路类型信息")
                    
                    for i, (lon, lat) in enumerate(result["all_coordinates"]):
                        road_type = ""
                        road_name = ""
                        
                        if "road_types" in result and i < len(result["road_types"]):
                            road_type_code = result["road_types"][i]
                            # 调试：打印当前点的道路类型
                            if i % 100 == 0:  # 每100个点打印一次，避免输出过多
                                print(f"点{i} 道路类型码: '{road_type_code}'")
                            
                            # 确保road_type_code是字符串
                            road_type_code = str(road_type_code).strip()
                            
                            if road_type_code == "1":
                                road_type = "高速公路"
                            elif road_type_code == "2":
                                road_type = "城市高架"
                            else:
                                road_type = "普通道路"
                        
                        # 获取道路名称
                        if "road_names" in result and i < len(result["road_names"]):
                            road_name = result["road_names"][i]
                        
                        coords_data.append({
                            "经度": lon,
                            "纬度": lat,
                            "道路类型": road_type,
                            "道路名称": road_name
                        })
                    
                    coords_df = pd.DataFrame(coords_data)
                    coords_df.index = coords_df.index + 1
                    coords_df.to_excel(writer, sheet_name='所有坐标点')
                    
                    # 导出路段信息
                    segments_df = pd.DataFrame(result["route_segments"])
                    segments_df.to_excel(writer, sheet_name='路段信息')
                    
                    # 导出各路段详细信息
                    start_idx = 0
                    for i, segment in enumerate(result["route_segments"]):
                        segment_coords = result["all_coordinates"][start_idx:start_idx + segment['坐标点数']]
                        segment_road_types = []
                        segment_road_names = []
                        
                        # 获取该路段的道路类型和名称
                        if "road_types" in result:
                            segment_road_types = result["road_types"][start_idx:start_idx + segment['坐标点数']]
                            # 调试：打印当前路段的道路类型统计
                            seg_type_counts = {}
                            for rt in segment_road_types:
                                seg_type_counts[rt] = seg_type_counts.get(rt, 0) + 1
                            print(f"路段{i+1} 道路类型统计:")
                            for rt, count in seg_type_counts.items():
                                print(f"  类型 '{rt}': {count}个点")
                        
                        if "road_names" in result:
                            segment_road_names = result["road_names"][start_idx:start_idx + segment['坐标点数']]
                        
                        # 创建路段数据
                        segment_data = []
                        for j, (lon, lat) in enumerate(segment_coords):
                            road_type = ""
                            road_name = ""
                            
                            if segment_road_types and j < len(segment_road_types):
                                road_type_code = str(segment_road_types[j]).strip()
                                if road_type_code == "1":
                                    road_type = "高速公路"
                                elif road_type_code == "2":
                                    road_type = "城市高架"
                                else:
                                    road_type = "普通道路"
                            
                            if segment_road_names and j < len(segment_road_names):
                                road_name = segment_road_names[j]
                            
                            segment_data.append({
                                "经度": lon,
                                "纬度": lat,
                                "道路类型": road_type,
                                "道路名称": road_name
                            })
                        
                        segment_df = pd.DataFrame(segment_data)
                        segment_df.index = segment_df.index + 1
                        sheet_name = f"路段{i+1}"
                        segment_df.to_excel(writer, sheet_name=sheet_name)
                        start_idx += segment['坐标点数']
                    
                    # 添加统计信息表
                    # 计算高速和高架里程
                    total_distance = 0
                    highway_distance = 0
                    elevated_distance = 0
                    
                    # 路段距离和类型统计
                    segment_distances = []  # 存储每段距离
                    segment_types = []      # 存储每段类型
                    segment_names = []      # 存储每段道路名称
                    
                    if "road_types" in result:
                        for i in range(len(result["all_coordinates"]) - 1):
                            if i < len(result["road_types"]):
                                road_type = str(result["road_types"][i]).strip()
                                road_name = ""
                                if "road_names" in result and i < len(result["road_names"]):
                                    road_name = result["road_names"][i]
                                
                                lon1, lat1 = result["all_coordinates"][i]
                                lon2, lat2 = result["all_coordinates"][i + 1]
                                
                                # 计算两点间距离
                                distance = self.calculate_distance(lat1, lon1, lat2, lon2) / 1000  # 转换为公里
                                total_distance += distance
                                
                                # 存储距离、类型和名称
                                segment_distances.append(distance)
                                segment_types.append(road_type)
                                segment_names.append(road_name)
                                
                                # 根据道路类型累加里程
                                if road_type == "1":  # 高速公路
                                    highway_distance += distance
                                elif road_type == "2":  # 城市高架
                                    elevated_distance += distance
                    
                    # 计算百分比
                    highway_percent = highway_distance / total_distance * 100 if total_distance > 0 else 0
                    elevated_percent = elevated_distance / total_distance * 100 if total_distance > 0 else 0
                    normal_percent = (total_distance - highway_distance - elevated_distance) / total_distance * 100 if total_distance > 0 else 0
                    
                    # 创建统计表
                    stats_data = [
                        {"类型": "总里程", "里程(公里)": round(total_distance, 2), "百分比(%)": 100.0},
                        {"类型": "高速公路", "里程(公里)": round(highway_distance, 2), "百分比(%)": round(highway_percent, 1)},
                        {"类型": "城市高架", "里程(公里)": round(elevated_distance, 2), "百分比(%)": round(elevated_percent, 1)},
                        {"类型": "普通道路", "里程(公里)": round(total_distance - highway_distance - elevated_distance, 2), "百分比(%)": round(normal_percent, 1)}
                    ]
                    stats_df = pd.DataFrame(stats_data)
                    stats_df.to_excel(writer, sheet_name='里程统计')
                    
                    # 添加路段类型详细统计表
                    if segment_distances and segment_types:
                        # 创建路段类型详细统计
                        segment_stats = []
                        for i, (distance, road_type, road_name) in enumerate(zip(segment_distances, segment_types, segment_names)):
                            road_type_name = ""
                            if road_type == "1":
                                road_type_name = "高速公路"
                            elif road_type == "2":
                                road_type_name = "城市高架"
                            else:
                                road_type_name = "普通道路"
                                
                            segment_stats.append({
                                "路段序号": i + 1,
                                "道路类型": road_type_name,
                                "道路类型码": road_type,
                                "道路名称": road_name,
                                "里程(公里)": round(distance, 4)
                            })
                        
                        segment_stats_df = pd.DataFrame(segment_stats)
                        segment_stats_df.to_excel(writer, sheet_name='路段类型详细统计')
                    
                self.log_text.append(f"已导出: {excel_path}")
            except Exception as e:
                self.log_text.append(f"{excel_path} 导出失败: {str(e)}")
        #QMessageBox.information(self, "导出成功", f"所有Excel已保存到: {output_dir}")
    
    def calculate_distance(self, lat1, lon1, lat2, lon2):
        """计算两点间的距离（单位：米）"""
        # 地球半径（米）
        R = 6371000
        
        # 将经纬度转换为弧度
        lat1_rad = math.radians(lat1)
        lon1_rad = math.radians(lon1)
        lat2_rad = math.radians(lat2)
        lon2_rad = math.radians(lon2)
        
        # 计算差值
        dlat = lat2_rad - lat1_rad
        dlon = lon2_rad - lon1_rad
        
        # Haversine公式
        a = math.sin(dlat/2)**2 + math.cos(lat1_rad) * math.cos(lat2_rad) * math.sin(dlon/2)**2
        c = 2 * math.atan2(math.sqrt(a), math.sqrt(1-a))
        distance = R * c
        
        return distance
    
    def select_excel_files(self):
        """选择Excel文件"""
        files, _ = QFileDialog.getOpenFileNames(self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls)")
        if files:
            for file in files:
                if file not in self.excel_files:
                    self.excel_files.append(file)
                    self.file_list.addItem(os.path.basename(file))
            
            self.map_log.append(f"已添加 {len(files)} 个Excel文件")
    
    def remove_selected_files(self):
        """移除选中的文件"""
        selected_items = self.file_list.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "提示", "请先选择要移除的文件")
            return
        
        for item in selected_items:
            row = self.file_list.row(item)
            del self.excel_files[row]
            self.file_list.takeItem(row)
        
        self.map_log.append(f"已移除 {len(selected_items)} 个文件")
    
    def clear_files(self):
        """清空文件列表"""
        self.excel_files = []
        self.file_list.clear()
        self.map_log.append("已清空文件列表")
    
    def generate_map(self):
        """生成地图"""
        if not self.excel_files:
            QMessageBox.warning(self, "警告", "请先选择至少一个Excel文件")
            return
        
        output_dir = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if not output_dir:
            return
        
        self.generate_html_map(output_dir)
    
    def generate_html_map(self, output_dir):
        """生成HTML地图"""
        self.map_progress.setValue(0)
        self.generate_btn.setEnabled(False)
        
        # 创建并启动生成线程
        self.gen_thread = RouteGenerator(self.excel_files, output_dir)
        self.gen_thread.progress_updated.connect(self.update_map_progress)
        self.gen_thread.generation_finished.connect(self.on_generation_finished)
        self.gen_thread.error_occurred.connect(self.on_generation_error)
        self.gen_thread.start()
    
    def update_map_progress(self, value):
        """更新地图生成进度"""
        self.map_progress.setValue(value)
    
    def on_generation_finished(self, html_path):
        """地图生成完成的回调"""
        self.generate_btn.setEnabled(True)
        self.map_log.append(f"地图已生成: {html_path}")
        QMessageBox.information(self, "生成成功", f"地图已保存到: {html_path}")
    
    def on_generation_error(self, error_msg):
        """地图生成错误的回调"""
        self.generate_btn.setEnabled(True)
        QMessageBox.critical(self, "生成错误", f"生成地图时出错: {error_msg}")
        self.map_log.append(f"错误: {error_msg}")

    def create_one_click_tab(self):
        """创建一键处理选项卡"""
        one_click_tab = QWidget()
        one_click_layout = QVBoxLayout(one_click_tab)
        
        # 导入链接Excel组
        links_group = QGroupBox("导入导航链接Excel")
        links_layout = QVBoxLayout()
        
        self.links_excel_path = QLineEdit()
        self.links_excel_path.setReadOnly(True)
        browse_layout = QHBoxLayout()
        browse_layout.addWidget(self.links_excel_path)
        self.browse_links_btn = QPushButton("浏览")
        self.browse_links_btn.clicked.connect(self.browse_links_excel)
        browse_layout.addWidget(self.browse_links_btn)
        self.browse_links_btn.setFixedHeight(36)
        links_layout.addLayout(browse_layout)
        
        # Sheet页选择
        self.sheet_selection_label = QLabel("Sheet页选择:")
        self.sheet_combo = QComboBox()
        self.sheet_combo.addItem("所有Sheet页")
        links_layout.addWidget(self.sheet_selection_label)
        links_layout.addWidget(self.sheet_combo)
        
        # 列选择
        self.link_column_label = QLabel("链接所在列 (默认为A列):")
        self.link_column_input = QLineEdit("A")
        links_layout.addWidget(self.link_column_label)
        links_layout.addWidget(self.link_column_input)
        
        # 添加自动识别链接列的复选框
        self.auto_detect_check = QCheckBox("自动识别所有Sheet页中的导航链接列")
        self.auto_detect_check.setChecked(True)
        links_layout.addWidget(self.auto_detect_check)
        
        links_group.setLayout(links_layout)
        one_click_layout.addWidget(links_group)
        
        # 输出设置组
        output_group = QGroupBox("输出设置")
        output_layout = QVBoxLayout()
        
        self.output_dir_path = QLineEdit()
        self.output_dir_path.setReadOnly(True)
        output_browse_layout = QHBoxLayout()
        output_browse_layout.addWidget(self.output_dir_path)
        self.browse_output_btn = QPushButton("浏览")
        self.browse_output_btn.clicked.connect(self.browse_output_dir)
        output_browse_layout.addWidget(self.browse_output_btn)
        self.browse_output_btn.setFixedHeight(36)
        output_layout.addLayout(output_browse_layout)
        
        output_group.setLayout(output_layout)
        one_click_layout.addWidget(output_group)
        
        # 一键处理按钮
        self.one_click_btn = QPushButton("开始一键处理")
        self.one_click_btn.clicked.connect(lambda: [self.one_click_process(), self.tab_widget.setCurrentIndex(1)])
        self.one_click_btn.setFixedHeight(52)
        one_click_layout.addWidget(self.one_click_btn)
        
        # 进度条
        self.one_click_progress = QProgressBar()
        self.one_click_progress.setRange(0, 100)
        self.one_click_progress.setValue(0)
        one_click_layout.addWidget(self.one_click_progress)
        
        # 一键处理日志
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout()
        
        self.one_click_log = QTextEdit()
        self.one_click_log.setReadOnly(True)
        log_layout.addWidget(self.one_click_log)
        
        log_group.setLayout(log_layout)
        one_click_layout.addWidget(log_group)
        
        self.tab_widget.addTab(one_click_tab, "一键处理")
    
    def browse_links_excel(self):
        """浏览选择包含导航链接的Excel文件并选择链接所在列"""
        file_path, _ = QFileDialog.getOpenFileName(self, "选择包含导航链接的Excel文件", "", "Excel文件 (*.xlsx *.xls)")
        if file_path:
            self.links_excel_path.setText(file_path)
            self.log_one_click(f"已选择导航链接Excel文件: {os.path.basename(file_path)}")
            
            # 读取Excel文件并获取所有sheet页
            try:
                # 添加文件读取错误处理
                try:
                    xl = pd.ExcelFile(file_path)
                except Exception as e:
                    logger.error(f"读取Excel文件失败: {str(e)}")
                    QMessageBox.critical(self, "文件读取错误", f"无法读取Excel文件: {str(e)}\n请检查文件是否损坏或格式正确。")
                    return
                
                # 更新sheet页下拉列表
                self.sheet_combo.clear()
                self.sheet_combo.addItem("所有Sheet页")
                for sheet_name in xl.sheet_names:
                    self.sheet_combo.addItem(sheet_name)
                
                # 关键词列表，用于自动检测导航链接列
                keywords = ['高德导航链接', '高德链接', '导航链接', '路线链接', '地图链接', '链接']
                
                # 导航链接正则表达式
                link_pattern = re.compile(r'https?://')
                
                # 自动检测最佳的导航链接列
                best_sheet = None
                best_col = None
                best_link_count = 0
                
                # 遍历所有sheet页和列，查找最佳导航链接列
                for sheet_name in xl.sheet_names:
                    # 读取当前sheet
                    df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
                    
                    # 遍历当前sheet的所有列
                    for col_idx in range(df.shape[1]):
                        # 提取当前列数据
                        col_data = df.iloc[:, col_idx].dropna().tolist()
                        
                        # 检查列中是否包含导航链接
                        link_count = 0
                        for cell in col_data:
                            if isinstance(cell, str) and link_pattern.match(cell):
                                link_count += 1
                        
                        # 找到链接数量最多的列
                        if link_count > best_link_count:
                            best_sheet = sheet_name
                            best_col = col_idx
                            best_link_count = link_count
                        
                        # 如果列名包含关键词，优先考虑
                        if link_count > 0:
                            # 检查第一行是否为列名
                            if df.shape[0] > 0:
                                col_name = df.iloc[0, col_idx] if pd.notna(df.iloc[0, col_idx]) else ""
                                col_text = str(col_name).strip().lower()
                                for keyword in keywords:
                                    if keyword.lower() in col_text:
                                        # 包含关键词的列，优先级更高
                                        best_sheet = sheet_name
                                        best_col = col_idx
                                        best_link_count = link_count
                                        break
                
                # 如果找到了最佳导航链接列，自动选择
                if best_sheet and best_col is not None:
                    # 选择对应的sheet页
                    sheet_index = self.sheet_combo.findText(best_sheet)
                    if sheet_index != -1:
                        self.sheet_combo.setCurrentIndex(sheet_index)
                    
                    # 将列索引转换为字母(A, B, C...)
                    col_letter = chr(65 + best_col)
                    self.link_column_input.setText(col_letter)
                    
                    # 取消自动识别选项，因为已经手动指定了列
                    self.auto_detect_check.setChecked(False)
                    
                    self.log_one_click(f"自动识别到导航链接所在位置: Sheet页 '{best_sheet}'，列 '{col_letter}' 列，包含 {best_link_count} 个链接")
                else:
                    # 如果没有找到明确的导航链接列，保持默认设置
                    self.log_one_click(f"已加载Excel文件，包含 {len(xl.sheet_names)} 个Sheet页")
                    self.log_one_click("未自动检测到明确的导航链接列，请手动选择")
                
                # 创建列选择对话框，显示所有sheet页的所有列（可选显示）
                dialog = QDialog(self)
                dialog.setWindowTitle("Excel文件列信息")
                layout = QVBoxLayout(dialog)
                
                layout.addWidget(QLabel("所有Sheet页的列信息（自动检测结果已填入）:"))
                
                # 创建一个QTreeWidget来显示sheet页和列的层级结构
                tree_widget = QTreeWidget(dialog)
                tree_widget.setHeaderLabels(["Sheet页/列", "描述"])
                
                # 遍历所有sheet页
                for sheet_name in xl.sheet_names:
                    sheet_item = QTreeWidgetItem([sheet_name, f"Sheet页: {sheet_name}"])
                    tree_widget.addTopLevelItem(sheet_item)
                    
                    # 读取当前sheet的列信息
                    df = pd.read_excel(xl, sheet_name=sheet_name, nrows=1)
                    columns = df.columns.tolist()
                    
                    # 遍历当前sheet的所有列
                    for col_idx, col_name in enumerate(columns):
                        # 检查是否为导航链接列
                        is_link_col = False
                        col_text = str(col_name).strip().lower()
                        for keyword in keywords:
                            if keyword.lower() in col_text:
                                is_link_col = True
                                break
                        
                        # 创建列项
                        col_letter = chr(65 + col_idx)
                        col_desc = f"列 {col_letter}: {col_name}"
                        
                        # 如果是最佳列，特别标记
                        if sheet_name == best_sheet and col_idx == best_col:
                            col_desc += " (自动选择)"
                        elif is_link_col:
                            col_desc += " (可能的导航链接列)"
                        
                        col_item = QTreeWidgetItem([f"{sheet_name}: {col_letter}", col_desc])
                        sheet_item.addChild(col_item)
                
                # 展开所有节点
                tree_widget.expandAll()
                layout.addWidget(tree_widget)
                
                # 确认按钮
                btn_layout = QHBoxLayout()
                ok_btn = QPushButton("确定")
                ok_btn.setFixedHeight(44)
                ok_btn.clicked.connect(dialog.accept)
                btn_layout.addWidget(ok_btn)
                layout.addLayout(btn_layout)
                
                # 显示对话框
                dialog.exec_()
            except Exception as e:
                self.log_one_click(f"读取Excel信息失败: {str(e)}")
                QMessageBox.warning(self, "警告", f"读取Excel信息失败: {str(e)}")
    
    def browse_output_dir(self):
        """浏览选择输出目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if dir_path:
            self.output_dir_path.setText(dir_path)
            self.log_one_click(f"已选择输出目录: {dir_path}")
    
    def log_one_click(self, message):
        """记录一键处理日志"""
        self.one_click_log.append(f"{pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}")
        logger.info(message)
        # 滚动到底部
        self.one_click_log.moveCursor(self.one_click_log.textCursor().End)
    
    def parse_amap_url(self, url):
        """解析高德导航链接，提取路线信息"""
        try:
            parsed = urlparse(url)
            if parsed.hostname not in ['ditu.amap.com', 'amap.com']:
                raise ValueError("不是有效的高德地图链接")
            
            params = parse_qs(parsed.query)
            
            # 提取起点、终点和途经点
            route_data = {
                "routeName": f"路线_{pd.Timestamp.now().strftime('%Y%m%d%H%M%S')}",
                "pointList": []
            }
            
            # 添加起点
            if 'from[lnglat]' in params and 'from[name]' in params:
                lng, lat = params['from[lnglat]'][0].split(',')
                route_data['pointList'].append({
                    "lat": float(lat),
                    "lon": float(lng),
                    "name": params['from[name]'][0],
                    "address": params['from[name]'][0]
                })
            
            # 添加途经点
            via_index = 0
            while f'via[{via_index}][lnglat]' in params and f'via[{via_index}][name]' in params:
                lng, lat = params[f'via[{via_index}][lnglat]'][0].split(',')
                route_data['pointList'].append({
                    "lat": float(lat),
                    "lon": float(lng),
                    "name": params[f'via[{via_index}][name]'][0],
                    "address": params[f'via[{via_index}][name]'][0]
                })
                via_index += 1
            
            # 添加终点
            if 'to[lnglat]' in params and 'to[name]' in params:
                lng, lat = params['to[lnglat]'][0].split(',')
                route_data['pointList'].append({
                    "lat": float(lat),
                    "lon": float(lng),
                    "name": params['to[name]'][0],
                    "address": params['to[name]'][0]
                })
            
            if len(route_data['pointList']) < 2:
                raise ValueError("链接中未找到足够的坐标点")
            
            return route_data
        except Exception as e:
            self.log_one_click(f"解析链接失败: {str(e)}")
            logger.error(f"解析链接失败: {str(e)}", exc_info=True)
            return None
    
    def import_links_excel(self, file_path, column=None):
        """从Excel导入导航链接，支持选择特定sheet页和列"""
        try:
            import pandas as pd
            import re
            
            # 导航链接正则表达式
            link_pattern = re.compile(r'https?://')
            
            # 读取Excel文件
            xl = pd.ExcelFile(file_path)
            all_links = []
            
            # 获取用户选择的sheet页
            selected_sheet = self.sheet_combo.currentText()
            
            # 获取用户选择的列
            user_col = self.link_column_input.text().strip()
            
            # 是否自动检测链接列
            auto_detect = self.auto_detect_check.isChecked()
            
            # 确定要处理的sheet页列表
            if selected_sheet == "所有Sheet页":
                sheets_to_process = xl.sheet_names
            else:
                sheets_to_process = [selected_sheet]
            
            # 遍历要处理的sheet页
            for sheet_name in sheets_to_process:
                self.log_one_click(f"正在处理sheet页: {sheet_name}")
                
                # 读取当前sheet
                df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
                
                # 确定要处理的列列表
                if auto_detect:
                    # 自动检测所有列
                    cols_to_process = range(df.shape[1])
                else:
                    # 用户指定的列
                    if user_col:
                        # 将列字母转换为索引
                        col_idx = ord(user_col.upper()) - ord('A')
                        if col_idx < 0 or col_idx >= df.shape[1]:
                            self.log_one_click(f"警告：指定的列 {user_col} 在sheet页 {sheet_name} 中不存在，跳过该sheet")
                            continue
                        cols_to_process = [col_idx]
                    else:
                        # 默认处理第一列
                        cols_to_process = [0]
                
                # 遍历要处理的列
                for col_idx in cols_to_process:
                    # 提取当前列数据
                    col_data = df.iloc[:, col_idx].dropna().tolist()
                    
                    # 检查列中是否包含导航链接
                    link_count = 0
                    links_in_col = []
                    
                    for cell in col_data:
                        if isinstance(cell, str) and link_pattern.match(cell):
                            link_count += 1
                            links_in_col.append(cell)
                    
                    # 如果是自动检测模式，需要满足链接数量条件
                    if auto_detect:
                        if link_count >= max(3, len(col_data) * 0.5):
                            self.log_one_click(f"在sheet页 {sheet_name} 的列 {chr(ord('A') + col_idx)} 中发现 {link_count} 个导航链接")
                            all_links.extend(links_in_col)
                    else:
                        # 手动指定列模式，直接提取所有链接
                        if link_count > 0:
                            self.log_one_click(f"在sheet页 {sheet_name} 的列 {chr(ord('A') + col_idx)} 中发现 {link_count} 个导航链接")
                            all_links.extend(links_in_col)
                        else:
                            self.log_one_click(f"在sheet页 {sheet_name} 的列 {chr(ord('A') + col_idx)} 中未发现导航链接")
            
            self.log_one_click(f"从Excel中读取到总计 {len(all_links)} 个导航链接")
            return all_links
        except Exception as e:
            self.log_one_click(f"导入Excel失败: {str(e)}")
            logger.error(f"导入Excel失败: {str(e)}", exc_info=True)
            return None
    
    def generate_json_from_links(self, links, output_dir):
        """从导航链接生成JSON文件"""
        json_files = []
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            self.log_one_click(f"创建输出目录: {output_dir}")
        
        for i, link in enumerate(links):
            # 更新进度
            progress = int((i / len(links)) * 30) + 20  # 20-50% 为JSON生成阶段
            self.one_click_progress.setValue(progress)
            
            self.log_one_click(f"正在解析链接 {i+1}/{len(links)}")
            route_data = self.parse_amap_url(link)
            
            if route_data:
                # 生成JSON文件名
                json_filename = f"route_{i+1}_{pd.Timestamp.now().strftime('%Y%m%d%H%M%S')}.json"
                json_path = os.path.join(output_dir, json_filename)
                
                # 保存JSON文件
                self.processed_files_count += 1
        
                # 定期清理缓存
                if self.processed_files_count % self.CACHE_CLEANUP_INTERVAL == 0:
                    self.log_one_click(f"已处理{self.processed_files_count}个文件，清理内存缓存...")
                    import gc
                    gc.collect()
                    self.log_one_click("内存缓存清理完成")
                    
                # 批量处理检查点
                if self.processed_files_count % self.BATCH_SIZE == 0:
                    self.log_one_click(f"已完成{self.processed_files_count}个路线处理，稍作停顿以优化性能...")
                    QApplication.processEvents()  # 处理未完成的UI事件
                    time.sleep(0.2)  # 短暂停顿，降低系统负载
# 原代码缩进若存在问题，根据上下文推测正确缩进，假设正确缩进层级如此
#with open(json_path, 'w', encoding='utf-8') as f:
            # 添加详细进度日志
                  # 使用enumerate替代index方法以处理重复链接
                    total_links = len(links)
                    self.log_one_click(f"正在处理路线 {i+1}/{total_links}: {os.path.basename(json_path)}")
                    self.progress_bar.setValue(int((i+1) / total_links * 100))
                    QApplication.processEvents()  # 刷新UI以显示进度
                  
                  # 确保输出目录存在并添加错误处理
                try:
                      os.makedirs(os.path.dirname(json_path), exist_ok=True)
                      self.log_one_click(f"已确保输出目录存在: {os.path.dirname(json_path)}")
                except PermissionError:
                      self.log_one_click(f"权限错误: 无法创建目录 {os.path.dirname(json_path)}")
                      QMessageBox.critical(self, "权限错误", f"无法创建目录: {os.path.dirname(json_path)}\n请检查目录权限。")
                      continue
                except Exception as e:
                      self.log_one_click(f"创建目录失败: {str(e)}")
                      QMessageBox.critical(self, "目录创建错误", f"创建输出目录时出错: {str(e)}")
                      continue
                  
                  # 尝试写入JSON文件并添加详细错误处理
                try:
                      with open(json_path, 'w', encoding='utf-8') as f:
                          json.dump([route_data], f, ensure_ascii=False, indent=2)
                      self.log_one_click(f"成功生成JSON文件: {os.path.basename(json_path)}")
                except PermissionError:
                      self.log_one_click(f"权限错误: 无法写入文件 {json_path}")
                      QMessageBox.critical(self, "权限错误", f"无法写入文件: {json_path}\n请检查文件权限或关闭正在使用该文件的程序。")
                      continue
                except Exception as e:
                      self.log_one_click(f"生成JSON文件失败: {str(e)}")
                      QMessageBox.critical(self, "文件生成错误", f"创建JSON文件时出错: {str(e)}")
                      continue
                
                json_files.append(json_path)
                self.log_one_click(f"已生成JSON文件: {json_filename}")
        else:
            self.log_one_click(f"链接 {i+1} 解析失败，跳过")
        
        return json_files
    
    def one_click_process(self):
        # 检查文件数量，超过阈值时显示警告
        excel_path = self.links_excel_path.text().strip()
        if not excel_path:
            QMessageBox.warning(self, "警告", "请先选择包含导航链接的Excel文件")
            return

        # 获取Excel文件中的链接数量
        try:
            link_column = self.link_column_input.text().strip().upper()
            if not link_column:
                QMessageBox.warning(self, "警告", "请先选择链接所在列")
                return
            
            # 读取用户选择的Sheet页
            selected_sheet = self.sheet_combo.currentText()
            xl = pd.ExcelFile(excel_path)
            
            # 确定要读取的Sheet页
            if selected_sheet == "所有Sheet页":
                # 如果选择了所有Sheet页，只检查第一个Sheet页的链接数量
                sheet_name = xl.sheet_names[0]
            else:
                sheet_name = selected_sheet
            
            # 读取指定Sheet页
            df = pd.read_excel(xl, sheet_name=sheet_name)
            
            # 正确转换列字母为索引，支持多字符列名（如AA, AB等）
            col_index = 0
            for char in link_column:
                col_index = col_index * 26 + (ord(char) - ord('A') + 1)
            col_index -= 1  # 转换为0-based索引
            
            # 检查列索引是否有效
            if col_index < 0 or col_index >= len(df.columns):
                QMessageBox.warning(self, "警告", f"无效的列选择: {link_column}，在Sheet页 '{sheet_name}' 中不存在该列")
                return
            
            # 计算链接数量
            links_count = len(df) - df.iloc[:, col_index].isna().sum()
            if links_count > self.MAX_CONCURRENT_FILES * self.BATCH_SIZE:
                reply = QMessageBox.warning(self,
                                           "文件数量警告",
                                           f"检测到{links_count}个导航链接，数量较多。\n"
                                           f"建议分批处理以避免性能问题。\n"
                                           "是否继续?",
                                           QMessageBox.Yes | QMessageBox.No,
                                           QMessageBox.No)
                if reply == QMessageBox.No:
                    self.log_one_click("用户取消了处理操作")
                    return
        except Exception as e:
            self.log_one_click(f"分析文件内容失败: {str(e)}")
            QMessageBox.warning(self, "警告", f"分析文件内容失败: {str(e)}")
            return
        """一键处理主函数"""
        try:
            # 检查输入
            links_excel_path = self.links_excel_path.text().strip()
            output_dir = self.output_dir_path.text().strip()
            link_column = self.link_column_input.text().strip() or 'A'
            
            if not links_excel_path:
                QMessageBox.warning(self, "警告", "请选择包含导航链接的Excel文件")
                return
            
            if not output_dir:
                QMessageBox.warning(self, "警告", "请选择输出目录")
                return
            
            # 重置进度条和日志
            self.one_click_progress.setValue(0)
            self.one_click_log.clear()
            self.log_one_click("===== 开始一键处理流程 ====")
            
            # 步骤1: 导入Excel中的导航链接 (0-20%)
            self.log_one_click("步骤1/5: 从Excel导入导航链接...")
            links = self.import_links_excel(links_excel_path, link_column)
            if not links:
                QMessageBox.warning(self, "警告", "未从Excel中读取到任何导航链接")
                return
            self.one_click_progress.setValue(20)
            
            # 步骤2: 解析链接并生成JSON文件 (20-50%)
            self.log_one_click("步骤2/5: 解析链接并生成JSON文件...")
            json_dir = os.path.join(output_dir, "json_files")
            json_files = self.generate_json_from_links(links, json_dir)
            if not json_files:
                QMessageBox.warning(self, "警告", "未生成任何JSON文件")
                return
            self.one_click_progress.setValue(50)
            
            # 步骤3: 计算路线并生成Excel文件 (50-80%)
            self.log_one_click("步骤3/5: 计算路线并生成Excel文件...")
            excel_dir = os.path.join(output_dir, "excel_file")
            if not os.path.exists(excel_dir):
                os.makedirs(excel_dir)
            self.auto_export_dir = excel_dir
            
            # 切换到路线生成选项卡并导入JSON文件
            self.tab_widget.setCurrentIndex(1)  # 假设路线生成是第2个选项卡
            self.json_files = json_files
            self.json_data_list = []
            self.route_list.clear()
            
            for json_path in json_files:
                try:
                    with open(json_path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    if isinstance(data, list) and len(data) > 0 and 'pointList' in data[0]:
                        route_data = data[0]
                        route_name = route_data.get('routeName', os.path.basename(json_path))
                        self.json_data_list.append(route_data)
                        self.route_list.addItem(f"{route_name} ({os.path.basename(json_path)})")
                except Exception as e:
                    self.log_one_click(f"加载JSON文件失败: {str(e)}")
            
            if not self.json_data_list:
                QMessageBox.warning(self, "警告", "无法加载生成的JSON文件")
                return
            
            # 计算路线
            self.calculate_btn.setEnabled(True)
            self.calculate_route()
            
            # 导出Excel
            self.export_all()
            self.one_click_progress.setValue(80)
            
            # 步骤4: 生成地图 (80-95%)
            self.log_one_click("步骤4/5: 生成地图...")
            self.tab_widget.setCurrentIndex(2)  # 假设地图生成是第3个选项卡
            
            # 获取所有生成的Excel文件
            self.excel_files = [os.path.join(excel_dir, f) for f in os.listdir(excel_dir) if f.endswith('.xlsx')]
            self.file_list.clear()
            for excel_file in self.excel_files:
                self.file_list.addItem(os.path.basename(excel_file))
            
            # 提示用户可以生成地图
            self.one_click_progress.setValue(95)
            #QMessageBox.information(self, "Excel导入成功", "Excel文件已批量导入，您可以点击生成地图按钮开始生成地图。")
            
            # 完成
            self.log_one_click("===== 一键处理流程完成 ====")
            self.log_one_click(f"所有结果已保存到: {output_dir}")
            self.one_click_progress.setValue(100)
            #QMessageBox.information(self, "成功", f"一键处理完成！\n所有结果已保存到: {output_dir}")
        except Exception as e:
            self.log_one_click(f"一键处理失败: {str(e)}")
            logger.error(f"一键处理失败: {str(e)}", exc_info=True)
            QMessageBox.critical(self, "错误", f"一键处理失败: {str(e)}")
    
    def remove_selected_json(self):
        """移除所选JSON文件"""
        selected_items = self.route_list.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "提示", "请先选择要移除的文件")
            return
        for item in selected_items:
            row = self.route_list.row(item)
            del self.json_files[row]
            del self.json_data_list[row]
            self.route_list.takeItem(row)
        self.import_status.setText(f"已导入 {len(self.json_files)} 个JSON文件")
        if not self.json_files:
            self.calculate_btn.setEnabled(False)
            self.export_btn.setEnabled(False)

    def clear_json_list(self):
        """清空JSON文件列表"""
        self.json_files = []
        self.json_data_list = []
        self.route_list.clear()
        self.import_status.setText("已清空JSON文件列表")
        self.calculate_btn.setEnabled(False)
        self.export_btn.setEnabled(False)

    def show_about_dialog(self):
        QMessageBox.information(
            self,
            "关于",
            f"泛化路线生成工具{VERSION}\n"
            "支持导入多个json文件，调用高德api生成导航路径；\n"
            "再导入导航路径excel生成html显示所有路线规划和里程。\n"
            "V1.1新增功能：计算高速和高架道路总里程，支持勾选显示隐藏路线。\n"
            "V1.2新增功能：增加可隐藏的路线图例，显示起终点和方向。\n"
            "V1.3新增功能：一键处理功能，支持导入包含导航链接的Excel文件，\n"
            "自动解析链接、生成JSON、生成坐标Excel并最终生成地图。\n"
            "更换高德地图作为底图，坐标系改为GCJ-02，\n"
            "去掉路线渲染颜色中的橘色和深橘色，避免与高德底图颜色重合。\n"
            "新增生成地图后自动打开html文件功能。\n"
            "设计by jiayu.tang"
        )

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # 设置应用程序字体
    font = QFont()
    font.setFamily("Microsoft YaHei")
    font.setPointSize(10)
    app.setFont(font)
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
