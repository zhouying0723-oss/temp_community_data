# -*- coding: utf-8 -*-
import os, json
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ========== 原始数据（2025-12-01 ~ 2026-07-06） ==========
raw_data = [
    # 2025年12月
    ("2025-12-01", 990914, 4054), ("2025-12-02", 1012612, 4254), ("2025-12-03", 1010840, 4373),
    ("2025-12-04", 1006405, 4042), ("2025-12-05", 981577, 3749), ("2025-12-06", 723677, 2550),
    ("2025-12-07", 560746, 2252), ("2025-12-08", 979486, 3783), ("2025-12-09", 1001761, 4081),
    ("2025-12-10", 1012719, 4218), ("2025-12-11", 1034741, 4788), ("2025-12-12", 994479, 4177),
    ("2025-12-13", 762590, 2844), ("2025-12-14", 571768, 2292), ("2025-12-15", 971913, 4154),
    ("2025-12-16", 1020486, 4667), ("2025-12-17", 1013829, 4344), ("2025-12-18", 1018149, 4680),
    ("2025-12-19", 1000679, 4295), ("2025-12-20", 727651, 2554), ("2025-12-21", 550088, 2352),
    ("2025-12-22", 974857, 4273), ("2025-12-23", 1016848, 4640), ("2025-12-24", 1012670, 4760),
    ("2025-12-25", 987369, 4259), ("2025-12-26", 975070, 4513), ("2025-12-27", 736971, 3320),
    ("2025-12-28", 582533, 2933), ("2025-12-29", 1022159, 5602), ("2025-12-30", 1032078, 5284),
    ("2025-12-31", 984399, 4537),
    # 2026年1月
    ("2026-01-01", 697434, 2671), ("2026-01-02", 738751, 3367), ("2026-01-03", 763515, 3456),
    ("2026-01-04", 777600, 3539), ("2026-01-05", 1050682, 4920), ("2026-01-06", 1064749, 5003),
    ("2026-01-07", 1064528, 5242), ("2026-01-08", 1055436, 5216), ("2026-01-09", 1049381, 5024),
    ("2026-01-10", 783241, 3538), ("2026-01-11", 618151, 3065), ("2026-01-12", 1077513, 5282),
    ("2026-01-13", 1107598, 5708), ("2026-01-14", 1116877, 5823), ("2026-01-15", 1117369, 5957),
    ("2026-01-16", 1061788, 5184), ("2026-01-17", 782239, 3332), ("2026-01-18", 639477, 3088),
    ("2026-01-19", 1039566, 4897), ("2026-01-20", 1080996, 5209), ("2026-01-21", 1097765, 5843),
    ("2026-01-22", 1090907, 5831), ("2026-01-23", 1081513, 5975), ("2026-01-24", 819534, 4080),
    ("2026-01-25", 648203, 3577), ("2026-01-26", 1089991, 6180), ("2026-01-27", 1110726, 6919),
    ("2026-01-28", 1160277, 7889), ("2026-01-29", 1228498, 10498), ("2026-01-30", 1262245, 12674),
    ("2026-01-31", 1061808, 9406),
    # 2026年2月
    ("2026-02-01", 834449, 6280), ("2026-02-02", 1317266, 12382), ("2026-02-03", 1276626, 9864),
    ("2026-02-04", 1268744, 9345), ("2026-02-05", 1262299, 9713), ("2026-02-06", 1228655, 8675),
    ("2026-02-07", 959856, 4773), ("2026-02-08", 704573, 3529), ("2026-02-09", 1148995, 6321),
    ("2026-02-10", 1148909, 6006), ("2026-02-11", 1136855, 5870), ("2026-02-12", 1108698, 5171),
    ("2026-02-13", 1087050, 5478), ("2026-02-14", 836545, 3099), ("2026-02-15", 628715, 2317),
    ("2026-02-16", 739366, 2711), ("2026-02-17", 785085, 3106), ("2026-02-18", 785266, 3110),
    ("2026-02-19", 781480, 3235), ("2026-02-20", 804540, 3451), ("2026-02-21", 790059, 3086),
    ("2026-02-22", 640467, 2391), ("2026-02-23", 903843, 4369), ("2026-02-24", 1162268, 5748),
    ("2026-02-25", 1167330, 5453), ("2026-02-26", 1174942, 5961), ("2026-02-27", 1183555, 5895),
    ("2026-02-28", 1161031, 6225),
    # 2026年3月
    ("2026-03-01", 1013315, 5611), ("2026-03-02", 1347844, 8387), ("2026-03-03", 1351006, 8262),
    ("2026-03-04", 1341788, 8010), ("2026-03-05", 1299978, 7139), ("2026-03-06", 1274324, 6838),
    ("2026-03-07", 1004620, 4472), ("2026-03-08", 783047, 3627), ("2026-03-09", 1287956, 7218),
    ("2026-03-10", 1294487, 6945), ("2026-03-11", 1288497, 6632), ("2026-03-12", 1285044, 6279),
    ("2026-03-13", 1268924, 6259), ("2026-03-14", 982524, 4376), ("2026-03-15", 801533, 3737),
    ("2026-03-16", 1267504, 6293), ("2026-03-17", 1270361, 5864), ("2026-03-18", 1286717, 6561),
    ("2026-03-19", 1347459, 8092), ("2026-03-20", 1305185, 6796), ("2026-03-21", 1005247, 4688),
    ("2026-03-22", 842322, 4072), ("2026-03-23", 1431359, 9629), ("2026-03-24", 1402543, 7913),
    ("2026-03-25", 1362207, 6915), ("2026-03-26", 1337357, 6432), ("2026-03-27", 1323689, 6018),
    ("2026-03-28", 1024829, 4035), ("2026-03-29", 810717, 3380), ("2026-03-30", 1294899, 5926),
    ("2026-03-31", 1305405, 5518),
    # 2026年4月
    ("2026-04-01", 1324127, 5994), ("2026-04-02", 1331661, 6306), ("2026-04-03", 1266079, 4905),
    ("2026-04-04", 857954, 2961), ("2026-04-05", 747474, 2839), ("2026-04-06", 1015920, 4259),
    ("2026-04-07", 1306284, 5524), ("2026-04-08", 1409776, 6730), ("2026-04-09", 1352125, 5922),
    ("2026-04-10", 1297181, 5038), ("2026-04-11", 987708, 3490), ("2026-04-12", 905167, 3873),
    ("2026-04-13", 1311431, 5479), ("2026-04-14", 1295368, 5018), ("2026-04-15", 1292279, 4893),
    ("2026-04-16", 1280175, 4656), ("2026-04-17", 1284058, 4877), ("2026-04-18", 1039217, 3582),
    ("2026-04-19", 817110, 3125), ("2026-04-20", 1272666, 4847), ("2026-04-21", 1284351, 4901),
    ("2026-04-22", 1270675, 4877), ("2026-04-23", 1267728, 5031), ("2026-04-24", 1263222, 5043),
    ("2026-04-25", 916309, 2871), ("2026-04-26", 768172, 2690), ("2026-04-27", 1208252, 4610),
    ("2026-04-28", 1240181, 4630), ("2026-04-29", 1234415, 4429), ("2026-04-30", 1202546, 4155),
    # 2026年5月
    ("2026-05-01", 906841, 2796), ("2026-05-02", 815379, 2263), ("2026-05-03", 670309, 1934),
    ("2026-05-04", 900993, 2981), ("2026-05-05", 998869, 3227), ("2026-05-06", 1256736, 4536),
    ("2026-05-07", 1250968, 4433), ("2026-05-08", 1248021, 4441), ("2026-05-09", 998482, 2637),
    ("2026-05-10", 734794, 2267), ("2026-05-11", 1230500, 4124), ("2026-05-12", 1244904, 4385),
    ("2026-05-13", 1284565, 4407), ("2026-05-14", 1295919, 4384), ("2026-05-15", 1286894, 4487),
    ("2026-05-16", 923563, 2699), ("2026-05-17", 747602, 2347), ("2026-05-18", 1234690, 4376),
    ("2026-05-19", 1236846, 4323), ("2026-05-20", 1254695, 4392), ("2026-05-21", 1267352, 4646),
    ("2026-05-22", 1235488, 4091), ("2026-05-23", 947186, 2554), ("2026-05-24", 768694, 2619),
    ("2026-05-25", 1201080, 3989), ("2026-05-26", 1227753, 4107), ("2026-05-27", 1249255, 4459),
    ("2026-05-28", 1238903, 4430), ("2026-05-29", 1225926, 4072), ("2026-05-30", 872072, 2226),
    ("2026-05-31", 663601, 1992),
    # 2026年6月
    ("2026-06-01", 1182600, 3801), ("2026-06-02", 1213861, 3961), ("2026-06-03", 1213703, 4152),
    ("2026-06-04", 1212854, 4168), ("2026-06-05", 1211218, 4334), ("2026-06-06", 965831, 3009),
    ("2026-06-07", 722395, 2354), ("2026-06-08", 1261710, 4822), ("2026-06-09", 1257023, 4419),
    ("2026-06-10", 1310748, 5763), ("2026-06-11", 1320599, 5317), ("2026-06-12", 1321505, 5104),
    ("2026-06-13", 986432, 2705), ("2026-06-14", 743602, 2309), ("2026-06-15", 1266008, 4509),
    ("2026-06-16", 1267815, 4527), ("2026-06-17", 1255257, 4303), ("2026-06-18", 1259098, 4537),
    ("2026-06-19", 1025799, 3262), ("2026-06-20", 829099, 2309), ("2026-06-21", 760663, 2373),
    ("2026-06-22", 1235126, 4103), ("2026-06-23", 1276774, 4475), ("2026-06-24", 1290074, 4518),
    ("2026-06-25", 1289016, 4801), ("2026-06-26", 1266258, 4397), ("2026-06-27", 961133, 2555),
    ("2026-06-28", 744261, 2356), ("2026-06-29", 1216927, 4112), ("2026-06-30", 1240069, 3928),
    # 2026年7月
    ("2026-07-01", 1257607, 4248), ("2026-07-02", 1278013, 4585), ("2026-07-03", 1237187, 3918),
    ("2026-07-04", 834422, 1985), ("2026-07-05", 651250, 1907), ("2026-07-06", 1196419, 3777),
]

dates_full = [d[0] for d in raw_data]
dau_full = [d[1] for d in raw_data]
community_full = [d[2] for d in raw_data]

# ========== 周均计算（每7天一周，不足7天丢弃） ==========
def calc_weekly(dates, dau_list, com_list):
    weeks = []
    week_labels = []
    for i in range(0, len(dates), 7):
        if i + 7 > len(dates):
            break
        week = dates[i:i+7]
        week_dau_avg = round(sum(dau_list[i:i+7]) / 7)
        week_com_avg = round(sum(com_list[i:i+7]) / 7)
        # 格式: 251201~251207
        sd = week[0][2:4] + week[0][5:7] + week[0][8:10]
        ed = week[6][5:7] + week[6][8:10]
        weeks.append({'dau': week_dau_avg, 'community': week_com_avg})
        week_labels.append(f"{sd}~{ed}")
    return week_labels, weeks

week_labels, weekly_data = calc_weekly(dates_full, dau_full, community_full)
weeks_dict = {}
for i, wl in enumerate(week_labels):
    weeks_dict[wl] = weekly_data[i]

# ========== 月均 ==========
monthly_dict = {}
for d, dd, dc in zip(dates_full, dau_full, community_full):
    m = d[:7]
    if m not in monthly_dict:
        monthly_dict[m] = {'dau_list': [], 'com_list': []}
    monthly_dict[m]['dau_list'].append(dd)
    monthly_dict[m]['com_list'].append(dc)

months_dict = {}
for m in sorted(monthly_dict.keys()):
    d_list = monthly_dict[m]['dau_list']
    c_list = monthly_dict[m]['com_list']
    months_dict[m] = {'dau': round(sum(d_list)/len(d_list)), 'community': round(sum(c_list)/len(c_list))}

save_dir = os.path.dirname(os.path.abspath(__file__))

# ========== 保存 Excel ==========
wb = openpyxl.Workbook()
header_font = Font(bold=True, size=11, color='FFFFFF')
header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
center_align = Alignment(horizontal='center', vertical='center')
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

def write_header(ws, headers):
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = header_font; c.fill = header_fill; c.alignment = center_align; c.border = thin_border

# Sheet1: 日数据
ws1 = wb.active; ws1.title = '日数据'
write_header(ws1, ['日期', '主动DAU', '社区首页UV'])
for i, (d, dau, cu) in enumerate(raw_data):
    ws1.cell(row=i+2, column=1, value=d).border = thin_border
    ws1.cell(row=i+2, column=1).alignment = center_align
    ws1.cell(row=i+2, column=2, value=dau).border = thin_border
    ws1.cell(row=i+2, column=3, value=cu).border = thin_border
ws1.column_dimensions['A'].width = 14; ws1.column_dimensions['B'].width = 14; ws1.column_dimensions['C'].width = 14

# Sheet2: 周均
ws2 = wb.create_sheet('周均')
write_header(ws2, ['周期', '主动DAU(周均)', '社区首页UV(周均)'])
for i, wl in enumerate(week_labels):
    ws2.cell(row=i+2, column=1, value=wl).border = thin_border
    ws2.cell(row=i+2, column=1).alignment = center_align
    ws2.cell(row=i+2, column=2, value=weekly_data[i]['dau']).border = thin_border
    ws2.cell(row=i+2, column=3, value=weekly_data[i]['community']).border = thin_border
ws2.column_dimensions['A'].width = 14; ws2.column_dimensions['B'].width = 16; ws2.column_dimensions['C'].width = 18

# Sheet3: 月均
ws3 = wb.create_sheet('月均')
write_header(ws3, ['月份', '主动DAU(月均)', '社区首页UV(月均)'])
for i, m in enumerate(sorted(months_dict.keys())):
    ws3.cell(row=i+2, column=1, value=m).border = thin_border
    ws3.cell(row=i+2, column=1).alignment = center_align
    ws3.cell(row=i+2, column=2, value=months_dict[m]['dau']).border = thin_border
    ws3.cell(row=i+2, column=3, value=months_dict[m]['community']).border = thin_border
ws3.column_dimensions['A'].width = 10; ws3.column_dimensions['B'].width = 16; ws3.column_dimensions['C'].width = 18

xlsx_path = os.path.join(save_dir, '主动DAU与社区首页UV数据.xlsx')
wb.save(xlsx_path)
print(f'Excel已保存: {xlsx_path}')

# ========== 计算功能渗透率 ==========
def calc_penetration(dau_val, com_val):
    """功能渗透率 = 社区首页UV / 主动DAU，返回百分比"""
    if dau_val == 0:
        return 0
    return round(com_val / dau_val * 100, 2)

# 日渗透率
days_penetration = {}
for d, dau, cu in zip(dates_full, dau_full, community_full):
    days_penetration[d] = round(cu / dau * 100, 2) if dau > 0 else 0

# 周渗透率(周均)
weeks_penetration = {}
for wl, wd in zip(week_labels, weekly_data):
    weeks_penetration[wl] = round(wd['community'] / wd['dau'] * 100, 2) if wd['dau'] > 0 else 0

# 月渗透率(月均)
months_penetration = {}
for m in sorted(months_dict.keys()):
    d = months_dict[m]
    months_penetration[m] = round(d['community'] / d['dau'] * 100, 2) if d['dau'] > 0 else 0

# ========== 日数据字典 ==========
days_dict = {}
for d, dau, cu in zip(dates_full, dau_full, community_full):
    days_dict[d] = {'dau': dau, 'community': cu}

# ========== 生成 ECharts HTML ==========
html = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>主动DAU & 社区首页UV</title>
<script src="https://cdn.jsdelivr.net/npm/echarts@5/dist/echarts.min.js"></script>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { background: #f5f7fa; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; padding: 20px; }
.container { max-width: 1100px; margin: 0 auto; background: #fff; border-radius: 12px; padding: 30px; }
.chart-section { margin-bottom: 10px; }
h2 { text-align: center; color: #333; font-weight: 500; margin-bottom: 20px; }
h3 { text-align: center; color: #555; font-weight: 500; margin: 20px 0 15px 0; font-size: 16px; }
.tabs { display: flex; justify-content: center; gap: 10px; margin-bottom: 20px; }
.tab-btn { padding: 8px 24px; border: 1px solid #d0d5dd; border-radius: 6px; background: #fff; cursor: pointer; font-size: 14px; color: #555; transition: all .2s; }
.tab-btn:hover { border-color: #5470C6; color: #5470C6; }
.tab-btn.active { background: #5470C6; color: #fff; border-color: #5470C6; }
#chart { width: 100%; height: 420px; }
#chart2 { width: 100%; height: 350px; }
.table-wrap { margin-top: 20px; overflow-x: auto; max-height: 350px; overflow-y: auto; }
.separator { border: none; border-top: 1px solid #e8e8e8; margin: 25px 0; }
table { width: 100%; border-collapse: collapse; font-size: 12px; }
th { background: #f0f4ff; color: #333; font-weight: 600; padding: 8px 10px; border: 1px solid #e0e5ee; text-align: center; position: sticky; top: 0; z-index: 1; }
td { padding: 6px 10px; border: 1px solid #e0e5ee; text-align: center; color: #444; }
tr:nth-child(even) { background: #fafbfc; }
tr:hover { background: #f0f4ff; }
.info { text-align: center; color: #999; font-size: 12px; margin-top: 12px; }
</style>
</head>
<body>
<div class="container">
  <div class="tabs">
    <button class="tab-btn active" data-tab="daily">日数据</button>
    <button class="tab-btn" data-tab="weekly">周均</button>
    <button class="tab-btn" data-tab="monthly">月均</button>
  </div>

  <h2>DAU & 社区首页UV</h2>
  <div id="chart"></div>
  <div class="table-wrap">
    <table id="data-table">
      <thead><tr><th>日期</th><th>主动DAU</th><th>社区首页UV</th></tr></thead>
      <tbody></tbody>
    </table>
  </div>

  <hr class="separator">

  <h3>社区首页功能渗透率</h3>
  <div id="chart2"></div>
  <div class="table-wrap">
    <table id="penetration-table">
      <thead><tr><th>日期</th><th>功能渗透率</th></tr></thead>
      <tbody></tbody>
    </table>
  </div>

  <div class="info">数据周期: 2025-12-01 ~ 2026-07-06</div>
</div>
<script>
var daysData = ''' + json.dumps(days_dict) + ''';
var weeksData = ''' + json.dumps(weeks_dict) + ''';
var monthsData = ''' + json.dumps(months_dict) + ''';

var daysPenetration = ''' + json.dumps(days_penetration) + ''';
var weeksPenetration = ''' + json.dumps(weeks_penetration) + ''';
var monthsPenetration = ''' + json.dumps(months_penetration) + ''';

function buildChart(data, type) {
  var labels = Object.keys(data);
  var dauVals = labels.map(function(k) { return data[k].dau; });
  var comVals = labels.map(function(k) { return data[k].community; });

  var option = {
    tooltip: {
      trigger: 'axis',
      axisPointer: { type: 'cross' },
      backgroundColor: 'rgba(255,255,255,0.95)',
      borderColor: '#e0e0e0', borderWidth: 1, padding: [10, 14],
      formatter: function(params) {
        var s = '<div style="font-weight:600;margin-bottom:6px;">' + params[0].axisValue + '</div>';
        params.forEach(function(p) {
          var v = p.seriesIndex === 0 ? Number(p.value).toLocaleString() : p.value;
          s += '<div style="display:flex;align-items:center;gap:6px;margin:3px 0;">' +
            '<span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:' + p.color + ';"></span>' +
            p.seriesName + ': <strong>' + v + '</strong></div>';
        });
        return s;
      }
    },
    legend: { data: ['主动DAU', '社区首页UV'], top: 5, textStyle: { fontSize: 13 } },
    grid: { left: 60, right: 60, bottom: 40, top: 50 },
    xAxis: {
      type: 'category', data: labels,
      axisLabel: { fontSize: 10, rotate: type === 'daily' ? 45 : 0, interval: type === 'daily' ? 13 : 0 },
      axisLine: { lineStyle: { color: '#ccc' } }
    },
    yAxis: [
      {
        type: 'value', name: '主动DAU',
        nameTextStyle: { fontSize: 13, fontWeight: 'bold', padding: [0,0,0,40] },
        min: 0,
        splitLine: { lineStyle: { color: '#f0f0f0', type: 'dashed' } },
        axisLabel: { fontSize: 11, formatter: function(v) { return v >= 10000 ? (v/10000).toFixed(0) + '万' : v; } }
      },
      {
        type: 'value', name: '社区首页UV',
        nameTextStyle: { fontSize: 13, fontWeight: 'bold', padding: [0,40,0,0] },
        min: 0,
        splitLine: { show: false },
        axisLabel: { fontSize: 11 }
      }
    ],
    series: [
      {
        name: '主动DAU', type: 'line', yAxisIndex: 0,
        data: dauVals, smooth: true,
        symbol: 'none',
        lineStyle: { width: 2, color: '#5470C6' },
        itemStyle: { color: '#5470C6' },
        areaStyle: {
          color: { type: 'linear', x: 0, y: 0, x2: 0, y2: 1,
            colorStops: [
              { offset: 0, color: 'rgba(84,112,198,0.25)' },
              { offset: 1, color: 'rgba(84,112,198,0.02)' }
            ]
          }
        }
      },
      {
        name: '社区首页UV', type: 'line', yAxisIndex: 1,
        data: comVals, smooth: true,
        symbol: 'none',
        lineStyle: { width: 2, color: '#EE6666' },
        itemStyle: { color: '#EE6666' },
        areaStyle: {
          color: { type: 'linear', x: 0, y: 0, x2: 0, y2: 1,
            colorStops: [
              { offset: 0, color: 'rgba(238,102,102,0.25)' },
              { offset: 1, color: 'rgba(238,102,102,0.02)' }
            ]
          }
        }
      }
    ]
  };
  return option;
}

function buildPenetrationChart(data, type) {
  var labels = Object.keys(data);
  var vals = labels.map(function(k) { return data[k]; });

  var option = {
    tooltip: {
      trigger: 'axis',
      axisPointer: { type: 'cross' },
      backgroundColor: 'rgba(255,255,255,0.95)',
      borderColor: '#e0e0e0', borderWidth: 1, padding: [10, 14],
      formatter: function(params) {
        var p = params[0];
        return '<div style="font-weight:600;margin-bottom:6px;">' + p.axisValue + '</div>' +
          '功能渗透率: <strong>' + p.value + '%</strong>';
      }
    },
    grid: { left: 60, right: 30, bottom: 40, top: 30 },
    xAxis: {
      type: 'category', data: labels,
      axisLabel: { fontSize: 10, rotate: type === 'daily' ? 45 : 0, interval: type === 'daily' ? 13 : 0 },
      axisLine: { lineStyle: { color: '#ccc' } }
    },
    yAxis: {
      type: 'value', name: '渗透率(%)',
      nameTextStyle: { fontSize: 12, fontWeight: 'bold' },
      splitLine: { lineStyle: { color: '#f0f0f0', type: 'dashed' } },
      axisLabel: { fontSize: 11, formatter: '{value}%' }
    },
    series: [
      {
        name: '功能渗透率', type: 'line',
        data: vals, smooth: true,
        symbol: type === 'daily' ? 'none' : 'circle',
        symbolSize: 6,
        lineStyle: { width: 3, color: '#91CC75' },
        itemStyle: { color: '#91CC75' },
        areaStyle: {
          color: { type: 'linear', x: 0, y: 0, x2: 0, y2: 1,
            colorStops: [
              { offset: 0, color: 'rgba(145,204,117,0.35)' },
              { offset: 1, color: 'rgba(145,204,117,0.02)' }
            ]
          }
        }
      }
    ]
  };
  return option;
}

function renderTable(data, tab, tableId) {
  var labels = Object.keys(data);
  var tbody = document.querySelector('#' + tableId + ' tbody');
  tbody.innerHTML = '';
  labels.forEach(function(k) {
    var row = document.createElement('tr');
    var td1 = document.createElement('td'); td1.textContent = k;
    var val = data[k];
    if (typeof val === 'object') {
      var td2 = document.createElement('td'); td2.textContent = Number(val.dau).toLocaleString();
      var td3 = document.createElement('td'); td3.textContent = val.community;
      row.appendChild(td1); row.appendChild(td2); row.appendChild(td3);
    } else {
      var td2 = document.createElement('td'); td2.textContent = val + '%';
      row.appendChild(td1); row.appendChild(td2);
    }
    tbody.appendChild(row);
  });
}

var chartDom = document.getElementById('chart');
var myChart = echarts.init(chartDom);
var chartDom2 = document.getElementById('chart2');
var myChart2 = echarts.init(chartDom2);

myChart.setOption(buildChart(daysData, 'daily'));
myChart2.setOption(buildPenetrationChart(daysPenetration, 'daily'));
renderTable(daysData, 'daily', 'data-table');
renderTable(daysPenetration, 'daily', 'penetration-table');

document.querySelectorAll('.tab-btn').forEach(function(btn) {
  btn.addEventListener('click', function() {
    document.querySelectorAll('.tab-btn').forEach(function(b) { b.classList.remove('active'); });
    btn.classList.add('active');
    var tab = btn.getAttribute('data-tab');
    var data, penData, type;
    if (tab === 'daily') { data = daysData; penData = daysPenetration; type = 'daily'; }
    else if (tab === 'weekly') { data = weeksData; penData = weeksPenetration; type = 'weekly'; }
    else { data = monthsData; penData = monthsPenetration; type = 'monthly'; }

    var thead1 = document.querySelector('#data-table thead tr');
    var thead2 = document.querySelector('#penetration-table thead tr');
    if (tab === 'daily') {
      thead1.innerHTML = '<th>日期</th><th>主动DAU</th><th>社区首页UV</th>';
      thead2.innerHTML = '<th>日期</th><th>功能渗透率</th>';
    } else if (tab === 'weekly') {
      thead1.innerHTML = '<th>周期</th><th>主动DAU(周均)</th><th>社区首页UV(周均)</th>';
      thead2.innerHTML = '<th>周期</th><th>功能渗透率(周均)</th>';
    } else {
      thead1.innerHTML = '<th>月份</th><th>主动DAU(月均)</th><th>社区首页UV(月均)</th>';
      thead2.innerHTML = '<th>月份</th><th>功能渗透率(月均)</th>';
    }

    myChart.setOption(buildChart(data, type));
    myChart2.setOption(buildPenetrationChart(penData, type));
    renderTable(data, tab, 'data-table');
    renderTable(penData, tab, 'penetration-table');
  });
});

window.addEventListener('resize', function() { myChart.resize(); myChart2.resize(); });
</script>
</body>
</html>'''

html_path = os.path.join(save_dir, '主动DAU与社区首页UV_折线图.html')
with open(html_path, 'w', encoding='utf-8') as f:
    f.write(html)
print(f'HTML已生成: {html_path}')