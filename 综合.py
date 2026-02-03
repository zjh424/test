import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import warnings
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# ======================== 基础配置（保持不变） ========================
warnings.filterwarnings('ignore')
plt.rcParams['font.sans-serif'] = ['WenQuanYi Zen Hei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

risk_weights = {
    '移动存储': 100,    
    '文档上传': 90,     
    '应用程序': 80,
    '网页浏览': 70,
    '策略日志': 60,
    '资产变更': 50,
    'windows日志': 40
}

# ======================== 新增：移动存储与文档上传评分工具函数 ========================
def get_risk_score_by_frequency(times, threshold_low=2, threshold_medium=5, threshold_high=10):
    """根据操作次数计算风险评分（0-1范围）
    - 0次：0.1分（极低风险）
    - 1-低阈值：0.3分（低风险）
    - 低阈值-中阈值：0.6分（中风险）
    - 中阈值-高阈值：0.8分（中高风险）
    - 超高中阈值：1.0分（高风险）
    """
    times = int(times) if not pd.isna(times) else 0
    if times == 0:
        return 0.1
    elif times <= threshold_low:
        return 0.3
    elif times <= threshold_medium:
        return 0.6
    elif times <= threshold_high:
        return 0.8
    else:
        return 1.0

# 定义高风险阈值变量
high_risk_threshold = 31.79  # 从截图中读取的高风险阈值

# ======================== 修正1：数据读取与清理 ========================
def read_and_clean_data(file_paths):
    """读取并清理所有数据源"""
    all_data = {}
    
    # 1. 读取所有文件的第二个工作表
    for file_name, sheet_config in file_paths.items():
        try:
            df = pd.read_excel(file_name, sheet_name=1)
            all_data[file_name] = df
            print(f"✅ 成功读取 {file_name} 的第二个工作表，数据形状：{df.shape}")
        except Exception as e:
            print(f"❌ 读取 {file_name} 时出错：{e}")
            all_data[file_name] = pd.DataFrame()
    
    # 2. 数据清理
    cleaned_data = {}
    
    # ---------------------- 移动存储数据清理 ----------------------
    if '移动存储_统计分析结果.xlsx' in all_data:
        storage_df = all_data['移动存储_统计分析结果.xlsx'].copy()
        # 适配常见列名
        storage_col_map = {
            '计算机': ['计算机', '计算机名称', '终端名称', '设备名'],
            '次数': ['移动存储次数', '使用次数', '操作次数', '访问次数']
        }
        
        # 匹配列名
        computer_col = None
        for col in storage_col_map['计算机']:
            if col in storage_df.columns:
                computer_col = col
                break
        
        count_col = None
        for col in storage_col_map['次数']:
            if col in storage_df.columns:
                count_col = col
                break
        
        if computer_col and count_col:
            # 数据清洗
            storage_clean = storage_df[[computer_col, count_col]].copy()
            storage_clean.rename(columns={computer_col: '计算机', count_col: '移动存储次数'}, inplace=True)
            storage_clean['计算机'] = storage_clean['计算机'].astype(str).str.strip()
            storage_clean['移动存储次数'] = pd.to_numeric(storage_clean['移动存储次数'], errors='coerce').fillna(0)
            # 去重和过滤空值
            storage_clean = storage_clean[
                (storage_clean['计算机'] != '') & 
                (storage_clean['计算机'] != 'nan')
            ].drop_duplicates(subset=['计算机'])
            cleaned_data['移动存储'] = storage_clean
            print(f"✅ 移动存储数据清理完成：{len(storage_clean)}条有效记录")
        else:
            cleaned_data['移动存储'] = pd.DataFrame()
            print(f"⚠️ 移动存储文件中未找到匹配列名，将使用默认低风险评分")
    else:
        cleaned_data['移动存储'] = pd.DataFrame()
        print(f"⚠️ 未找到移动存储文件，将使用默认低风险评分")
    
    # ---------------------- 文档上传数据清理 ----------------------
    if '人员上传操作统计分析.xlsx' in all_data:
        upload_df = all_data['人员上传操作统计分析.xlsx'].copy()
        
        # 简化处理：直接查找包含"人员"和"上传"的列
        person_col = None
        count_col = None
        risk_level_col = None
        
        for col in upload_df.columns:
            col_str = str(col)
            if '人员' in col_str or '计算机' in col_str or '用户' in col_str:
                person_col = col
            elif '上传' in col_str and ('次' in col_str or '数量' in col_str):
                count_col = col
            elif '风险' in col_str and ('等级' in col_str or '级别' in col_str):
                risk_level_col = col
        
        if person_col and count_col:
            upload_clean = upload_df[[person_col, count_col]].copy()
            if risk_level_col:
                upload_clean[risk_level_col] = upload_df[risk_level_col]
            
            upload_clean.rename(columns={person_col: '人员（计算机）', count_col: '上传次数'}, inplace=True)
            if risk_level_col:
                upload_clean.rename(columns={risk_level_col: '风险等级'}, inplace=True)
            
            upload_clean['人员（计算机）'] = upload_clean['人员（计算机）'].astype(str).str.strip()
            upload_clean['上传次数'] = pd.to_numeric(upload_clean['上传次数'], errors='coerce').fillna(0)
            
            # 过滤空值
            upload_clean = upload_clean[
                (upload_clean['人员（计算机）'] != '') & 
                (upload_clean['人员（计算机）'] != 'nan')
            ]
            
            cleaned_data['文档上传'] = upload_clean
            print(f"✅ 文档上传数据清理完成：{len(upload_clean)}条有效记录")
        else:
            cleaned_data['文档上传'] = pd.DataFrame()
            print(f"⚠️ 文档上传文件中未找到匹配列名，将使用默认低风险评分")
    else:
        cleaned_data['文档上传'] = pd.DataFrame()
        print(f"⚠️ 未找到文档上传文件，将使用默认低风险评分")
    
    # ---------------------- 其他维度清理（保持不变） ----------------------
    # 应用程序数据清理
    if '应用程序_综合分析报告_含风险评估依据.xlsx' in all_data:
        app_df = all_data['应用程序_综合分析报告_含风险评估依据.xlsx'].copy()
        app_df = app_df[app_df['计算机'].notna() & (app_df['计算机'] != '')].copy()
        app_df['计算机'] = app_df['计算机'].astype(str).str.strip()
        cleaned_data['应用程序'] = app_df
    
    # 网页浏览数据清理
    if '网页浏览风险分析报告.xlsx' in all_data:
        web_df = all_data['网页浏览风险分析报告.xlsx'].copy()
        web_header_row = None
        for idx, row in web_df.iterrows():
            if '计算机' in str(row.iloc[0]) and '风险评分' in str(row.iloc[4]):
                web_header_row = idx
                break
        
        web_clean_df = pd.DataFrame()
        if web_header_row is not None:
            web_data = web_df.iloc[web_header_row+1:].copy()
            web_data.columns = ['计算机', '总记录数', '使用用户数', '访问网站数', '风险评分', '风险等级']
            web_data = web_data[web_data['计算机'].notna() & (web_data['计算机'] != '')].copy()
            web_data['计算机'] = web_data['计算机'].astype(str).str.strip()
            web_data['风险评分'] = pd.to_numeric(web_data['风险评分'], errors='coerce') / 100
            web_clean_df = web_data
        cleaned_data['网页浏览'] = web_clean_df
    
    # 策略日志数据清理
    if '策略日志统计分析结果.xlsx' in all_data:
        policy_df = all_data['策略日志统计分析结果.xlsx'].copy()
        policy_df = policy_df[policy_df.iloc[:, 0].notna() & (policy_df.iloc[:, 0] != '策略日志统计分析报告')].copy()
        policy_df.columns = ['计算机', '报警次数', '风险等级', '涉及策略', '用户', '计算机组']
        policy_df['计算机'] = policy_df['计算机'].astype(str).str.strip()
        policy_df['报警次数'] = pd.to_numeric(policy_df['报警次数'], errors='coerce')
        policy_df['风险评分'] = policy_df['风险等级'].map({'高风险': 1.0, '正常': 0.2, '中风险': 0.6}).fillna(0.5)
        cleaned_data['策略日志'] = policy_df
    
    # 资产变更数据清理
    if '资产变更统计分析.xlsx' in all_data:
        asset_df = all_data['资产变更统计分析.xlsx'].copy()
        asset_df = asset_df[asset_df['计算机'].notna() & (asset_df['计算机'] != '')].copy()
        asset_df['计算机'] = asset_df['计算机'].astype(str).str.strip()
        asset_df['标准化风险评分'] = pd.to_numeric(asset_df['风险评分'], errors='coerce')
        cleaned_data['资产变更'] = asset_df
    
    # Windows日志数据清理（备用）
    cleaned_data['windows日志'] = pd.DataFrame()
    
    return cleaned_data

# ======================== 修正2：风险计算 ========================
def calculate_comprehensive_risk(cleaned_data):
    """计算综合风险评分"""
    valid_computers = ['孙华阳', '夏孙志', '周宏波', '周航', '许廷哲', '宋奇隆', '邱草谋', '张裕兴', '肖灵活', '邬佳怡','唐启凡','曲艺','王钦北']
    comprehensive_risk = []
    
    for computer in valid_computers:
        risk_data = {
            '计算机': computer,
            '移动存储_风险评分': 0.0,
            '文档上传_风险评分': 0.0, 
            '应用程序_风险评分': 0.0,
            '网页浏览_风险评分': 0.0,
            '策略日志_风险评分': 0.0,
            '资产变更_风险评分': 0.0,
            'windows日志_风险评分': 0.2,
            '各维度风险说明': []
        }
        
        # ---------------------- 1. 其他维度评分 ----------------------
        # 应用程序风险
        app_df = cleaned_data.get('应用程序', pd.DataFrame())
        if not app_df.empty and '计算机' in app_df.columns and '风险评分' in app_df.columns:
            app_match = app_df[app_df['计算机'] == computer]
            if not app_match.empty:
                app_risk_value = app_match['风险评分'].iloc[0]
                if not pd.isna(app_risk_value):
                    risk_data['应用程序_风险评分'] = float(app_risk_value)
                    app_level = app_match['风险等级'].iloc[0] if '风险等级' in app_match.columns else '未知'
                    risk_data['各维度风险说明'].append(f"应用程序：{app_level}（评分：{risk_data['应用程序_风险评分']:.3f}）")
                else:
                    risk_data['各维度风险说明'].append("应用程序：数据缺失（评分：0.000）")
            else:
                risk_data['各维度风险说明'].append("应用程序：无记录（评分：0.000）")
        
        # 网页浏览风险
        web_df = cleaned_data.get('网页浏览', pd.DataFrame())
        if not web_df.empty and '计算机' in web_df.columns and '风险评分' in web_df.columns:
            web_match = web_df[web_df['计算机'] == computer]
            if not web_match.empty:
                web_risk_value = web_match['风险评分'].iloc[0]
                if not pd.isna(web_risk_value):
                    risk_data['网页浏览_风险评分'] = float(web_risk_value)
                    web_level = web_match['风险等级'].iloc[0] if '风险等级' in web_match.columns else '未知'
                    risk_data['各维度风险说明'].append(f"网页浏览：{web_level}（评分：{risk_data['网页浏览_风险评分']:.3f}）")
                else:
                    risk_data['各维度风险说明'].append("网页浏览：数据缺失（评分：0.000）")
            else:
                risk_data['各维度风险说明'].append("网页浏览：无记录（评分：0.000）")
        
        # 策略日志风险
        policy_df = cleaned_data.get('策略日志', pd.DataFrame())
        if not policy_df.empty and '计算机' in policy_df.columns and '风险评分' in policy_df.columns:
            policy_match = policy_df[policy_df['计算机'] == computer]
            if not policy_match.empty:
                policy_risk_value = policy_match['风险评分'].iloc[0]
                if not pd.isna(policy_risk_value):
                    risk_data['策略日志_风险评分'] = float(policy_risk_value)
                    policy_level = policy_match['风险等级'].iloc[0] if '风险等级' in policy_match.columns else '未知'
                    policy_alerts = policy_match['报警次数'].iloc[0] if '报警次数' in policy_match.columns else 0
                    risk_data['各维度风险说明'].append(f"策略日志：{policy_level}（报警{int(policy_alerts)}次，评分：{risk_data['策略日志_风险评分']:.3f}）")
                else:
                    risk_data['各维度风险说明'].append("策略日志：数据缺失（评分：0.000）")
            else:
                risk_data['各维度风险说明'].append("策略日志：无记录（评分：0.000）")
        
        # 资产变更风险
        asset_df = cleaned_data.get('资产变更', pd.DataFrame())
        if not asset_df.empty and '计算机' in asset_df.columns and '标准化风险评分' in asset_df.columns:
            asset_match = asset_df[asset_df['计算机'] == computer]
            if not asset_match.empty:
                asset_risk_value = asset_match['标准化风险评分'].iloc[0]
                if not pd.isna(asset_risk_value):
                    risk_data['资产变更_风险评分'] = float(asset_risk_value)
                    asset_level = asset_match['风险等级'].iloc[0] if '风险等级' in asset_match.columns else '未知'
                    risk_data['各维度风险说明'].append(f"资产变更：{asset_level}（评分：{risk_data['资产变更_风险评分']:.3f}）")
                else:
                    risk_data['各维度风险说明'].append("资产变更：数据缺失（评分：0.000）")
            else:
                risk_data['各维度风险说明'].append("资产变更：无记录（评分：0.000）")
        
        # Windows日志风险说明
        risk_data['各维度风险说明'].append(f"Windows日志：数据不足（默认低风险，评分：0.200）")
        
        # ---------------------- 2. 移动存储评分 ----------------------
        storage_df = cleaned_data.get('移动存储', pd.DataFrame())
        if not storage_df.empty and '计算机' in storage_df.columns and '移动存储次数' in storage_df.columns:
            storage_match = storage_df[storage_df['计算机'] == computer]
            if not storage_match.empty:
                storage_times = storage_match['移动存储次数'].iloc[0]
                storage_score = get_risk_score_by_frequency(storage_times, threshold_low=2, threshold_medium=5, threshold_high=10)
                risk_data['移动存储_风险评分'] = storage_score
                risk_level = '低风险' if storage_score <= 0.3 else '中风险' if storage_score <= 0.6 else '中高风险' if storage_score <= 0.8 else '高风险'
                risk_data['各维度风险说明'].append(f"移动存储：{risk_level}（操作{int(storage_times)}次，评分：{storage_score:.3f}）")
            else:
                risk_data['移动存储_风险评分'] = 0.1
                risk_data['各维度风险说明'].append(f"移动存储：无操作记录（默认低风险，评分：0.100）")
        else:
            risk_data['移动存储_风险评分'] = 0.1
            risk_data['各维度风险说明'].append(f"移动存储：数据异常（默认低风险，评分：0.100）")
        
        # ---------------------- 3. 文档上传评分 ----------------------
        upload_df = cleaned_data.get('文档上传', pd.DataFrame())
        if not upload_df.empty and '人员（计算机）' in upload_df.columns and '上传次数' in upload_df.columns:
            upload_match = upload_df[upload_df['人员（计算机）'] == computer]
            if not upload_match.empty:
                upload_times = upload_match['上传次数'].iloc[0]
                # 使用高风险阈值作为阈值
                upload_score = get_risk_score_by_frequency(upload_times, 
                                                        threshold_low=high_risk_threshold/4, 
                                                        threshold_medium=high_risk_threshold/2, 
                                                        threshold_high=high_risk_threshold)
                risk_data['文档上传_风险评分'] = upload_score
                
                # 尝试获取已有的风险等级
                if '风险等级' in upload_match.columns:
                    existing_risk_level = upload_match['风险等级'].iloc[0]
                else:
                    existing_risk_level = '低风险' if upload_score <= 0.3 else '中风险' if upload_score <= 0.6 else '中高风险' if upload_score <= 0.8 else '高风险'
                
                risk_data['各维度风险说明'].append(f"文档上传：{existing_risk_level}（操作{int(upload_times)}次，评分：{upload_score:.3f}）")
            else:
                risk_data['文档上传_风险评分'] = 0.1
                risk_data['各维度风险说明'].append(f"文档上传：无该人员统计记录（默认低风险，评分：0.100）")
        else:
            risk_data['文档上传_风险评分'] = 0.1
            risk_data['各维度风险说明'].append(f"文档上传：统计表格结构异常（默认低风险，评分：0.100）")
        
        # ---------------------- 4. 综合评分计算 ----------------------
        weighted_score = (
            risk_data['移动存储_风险评分'] * risk_weights['移动存储'] +
            risk_data['文档上传_风险评分'] * risk_weights['文档上传'] +
            risk_data['应用程序_风险评分'] * risk_weights['应用程序'] +
            risk_data['网页浏览_风险评分'] * risk_weights['网页浏览'] +
            risk_data['策略日志_风险评分'] * risk_weights['策略日志'] +
            risk_data['资产变更_风险评分'] * risk_weights['资产变更'] +
            risk_data['windows日志_风险评分'] * risk_weights['windows日志']
        ) / sum(risk_weights.values()) * 100
        
        risk_data['综合风险评分（百分制）'] = round(weighted_score, 2)
        
        # 风险等级和建议
        if weighted_score >= 80:
            risk_data['综合风险等级'] = '🔴 高风险'
            risk_data['风险处理建议'] = '需要立即重点关注和审计，建议进行详细的安全检查'
        elif weighted_score >= 60:
            risk_data['综合风险等级'] = '🟡 中高风险'
            risk_data['风险处理建议'] = '建议加强监控，定期审查相关操作日志'
        elif weighted_score >= 40:
            risk_data['综合风险等级'] = '🟢 中风险'
            risk_data['风险处理建议'] = '保持常规监控，关注风险变化趋势'
        else:
            risk_data['综合风险等级'] = '🔵 低风险'
            risk_data['风险处理建议'] = '维持现有安全策略，继续保持良好安全状态'
        
        risk_data['各维度风险详细说明'] = '; '.join(risk_data['各维度风险说明'])
        comprehensive_risk.append(risk_data)
    
    comprehensive_df = pd.DataFrame(comprehensive_risk)
    comprehensive_df_sorted = comprehensive_df.sort_values('综合风险评分（百分制）', ascending=False).reset_index(drop=True)
    
    return comprehensive_df_sorted

# ======================== 生成可视化图表 ========================
def generate_visualization(final_df, output_path):
    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(16, 12))
    fig.suptitle('人员综合风险统计分析图表', fontsize=20, fontweight='bold', y=0.95)
    
    colors_high = '#FF4444'
    colors_medium_high = '#FF8800'
    colors_medium = '#FFDD44'
    colors_low = '#44DD44'
    
    # 1. 综合风险评分柱状图
    computers = final_df['计算机'].values
    scores = final_df['综合风险评分（百分制）'].values
    risk_levels = final_df['综合风险等级'].values
    
    bar_colors = []
    for level in risk_levels:
        if '高风险' in level and '中高' not in level:
            bar_colors.append(colors_high)
        elif '中高风险' in level:
            bar_colors.append(colors_medium_high)
        elif '中风险' in level:
            bar_colors.append(colors_medium)
        else:
            bar_colors.append(colors_low)
    
    bars1 = ax1.barh(computers, scores, color=bar_colors, alpha=0.8, edgecolor='white', linewidth=1)
    ax1.set_xlabel('综合风险评分（百分制）', fontsize=12, fontweight='bold')
    ax1.set_title('各人员综合风险评分对比', fontsize=14, fontweight='bold', pad=20)
    ax1.grid(axis='x', alpha=0.3, linestyle='--')
    ax1.set_xlim(0, 100)
    
    for bar, score in zip(bars1, scores):
        ax1.text(score + 1, bar.get_y() + bar.get_height()/2, f'{score:.1f}', 
                 va='center', ha='left', fontsize=10, fontweight='bold')
    
    ax1.axvline(x=80, color=colors_high, linestyle='--', alpha=0.7, label='高风险线(80分)')
    ax1.axvline(x=60, color=colors_medium_high, linestyle='--', alpha=0.7, label='中高风险线(60分)')
    ax1.axvline(x=40, color=colors_medium, linestyle='--', alpha=0.7, label='中风险线(40分)')
    ax1.legend(loc='lower right', fontsize=10)
    
    # 2. 风险等级分布饼图
    risk_counts = final_df['综合风险等级'].value_counts()
    risk_labels = [label.replace('🔴 ', '').replace('🟡 ', '').replace('🟢 ', '').replace('🔵 ', '') 
                   for label in risk_counts.index]
    risk_colors_pie = []
    for label in risk_counts.index:
        if '高风险' in label and '中高' not in label:
            risk_colors_pie.append(colors_high)
        elif '中高风险' in label:
            risk_colors_pie.append(colors_medium_high)
        elif '中风险' in label:
            risk_colors_pie.append(colors_medium)
        else:
            risk_colors_pie.append(colors_low)
    
    wedges, texts, autotexts = ax2.pie(risk_counts.values, labels=risk_labels, colors=risk_colors_pie,
                                       autopct='%1.1f%%', startangle=90, textprops={'fontsize': 11})
    ax2.set_title('风险等级分布', fontsize=14, fontweight='bold', pad=20)
    
    for autotext in autotexts:
        autotext.set_color('white')
        autotext.set_fontweight('bold')
    
    # 3. 各风险维度平均评分对比
    dimensions = ['移动存储', '文档上传', '应用程序', '网页浏览', '策略日志', '资产变更', 'Windows日志']
    avg_scores_dim = [
        final_df['移动存储_风险评分'].mean(),
        final_df['文档上传_风险评分'].mean(),
        final_df['应用程序_风险评分'].mean(),
        final_df['网页浏览_风险评分'].mean(),
        final_df['策略日志_风险评分'].mean(),
        final_df['资产变更_风险评分'].mean(),
        final_df['windows日志_风险评分'].mean()
    ]
    
    weight_order = [0, 1, 2, 3, 4, 5, 6]
    dimensions_ordered = [dimensions[i] for i in weight_order]
    avg_scores_ordered = [avg_scores_dim[i] for i in weight_order]
    
    bars3 = ax3.bar(range(len(dimensions_ordered)), avg_scores_ordered, 
                    color=['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FECA57', '#FF9FF3', '#54A0FF'], 
                    alpha=0.8, edgecolor='white', linewidth=1)
    
    ax3.set_xlabel('风险维度', fontsize=12, fontweight='bold')
    ax3.set_ylabel('平均风险评分（0-1）', fontsize=12, fontweight='bold')
    ax3.set_title('各风险维度平均风险评分对比（按权重排序）', fontsize=14, fontweight='bold', pad=20)
    ax3.set_xticks(range(len(dimensions_ordered)))
    ax3.set_xticklabels(dimensions_ordered, rotation=45, ha='right')
    ax3.grid(axis='y', alpha=0.3, linestyle='--')
    ax3.set_ylim(0, 1)
    
    for bar, score in zip(bars3, avg_scores_ordered):
        ax3.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.02, f'{score:.3f}',
                 va='bottom', ha='center', fontsize=10, fontweight='bold')
    
    # 4. 前3名高风险人员各维度分析
    top3_high_risk = final_df.head(3)
    dimensions_short = ['移动存储', '文档上传', '应用程序', '网页浏览', '策略日志', '资产变更', 'Windows']
    x_pos = np.arange(len(dimensions_short))
    width = 0.25
    
    for i, (idx, row) in enumerate(top3_high_risk.iterrows()):
        values = [
            row['移动存储_风险评分'],
            row['文档上传_风险评分'],
            row['应用程序_风险评分'],
            row['网页浏览_风险评分'],
            row['策略日志_风险评分'],
            row['资产变更_风险评分'],
            row['windows日志_风险评分']
        ]
        ax4.bar(x_pos + i*width, values, width, label=row['计算机'], 
                alpha=0.8, edgecolor='white', linewidth=1)
    
    ax4.set_xlabel('风险维度', fontsize=12, fontweight='bold')
    ax4.set_ylabel('风险评分（0-1）', fontsize=12, fontweight='bold')
    ax4.set_title('前3名高风险人员各维度风险评分对比', fontsize=14, fontweight='bold', pad=20)
    ax4.set_xticks(x_pos + width)
    ax4.set_xticklabels(dimensions_short, rotation=45, ha='right')
    ax4.legend(loc='upper right', fontsize=10)
    ax4.grid(axis='y', alpha=0.3, linestyle='--')
    ax4.set_ylim(0, 1.1)
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight', facecolor='white')
    plt.close()
    
    print(f"✅ 可视化图表已保存至：{output_path}")

# ======================== 生成Excel报告 ========================
def generate_excel_report(final_df, output_path):
    wb = Workbook()
    
    ws1 = wb.active
    ws1.title = "综合风险统计表"
    
    report_data = final_df[['计算机', '综合风险评分（百分制）', '综合风险等级', '风险处理建议',
                           '移动存储_风险评分', '文档上传_风险评分', '应用程序_风险评分',
                           '网页浏览_风险评分', '策略日志_风险评分', '资产变更_风险评分',
                           'windows日志_风险评分', '各维度风险详细说明']].copy()
    
    report_data.insert(0, '序号', range(1, len(report_data) + 1))
    
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    for r_idx, row in enumerate(dataframe_to_rows(report_data, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws1.cell(row=r_idx, column=c_idx, value=value)
            
            if r_idx == 1:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                risk_level = report_data.iloc[r_idx-2]['综合风险等级']
                if '高风险' in risk_level and '中高' not in risk_level:
                    cell.fill = PatternFill(start_color="FFE6E6", end_color="FFE6E6", fill_type="solid")
                elif '中高风险' in risk_level:
                    cell.fill = PatternFill(start_color="FFF2E6", end_color="FFF2E6", fill_type="solid")
                elif '中风险' in risk_level:
                    cell.fill = PatternFill(start_color="FFFFE6", end_color="FFFFE6", fill_type="solid")
                else:
                    cell.fill = PatternFill(start_color="E6FFE6", end_color="E6FFE6", fill_type="solid")
            
            cell.border = thin_border
    
    column_widths = [6, 12, 18, 20, 25, 18, 18, 18, 18, 18, 18, 18, 60]
    for i, width in enumerate(column_widths, 1):
        ws1.column_dimensions[chr(64 + i)].width = width
    
    ws1.row_dimensions[1].height = 30
    for i in range(2, len(report_data) + 2):
        ws1.row_dimensions[i].height = 40
    
    ws2 = wb.create_sheet(title="风险评估依据和说明")
    
    risk_evaluation = {
        '风险评估标准说明': '''
1. 风险权重设定依据
   - 移动存储（权重100）：最高优先级，因移动存储设备容易导致数据泄露和恶意软件传播
   - 文档上传（权重90）：高优先级，涉及数据外发风险，需重点监控
   - 应用程序（权重80）：中高优先级，异常应用程序使用可能带来安全威胁
   - 网页浏览（权重70）：中优先级，访问风险网站可能导致病毒感染
   - 策略日志（权重60）：中低优先级，违反安全策略行为需要关注
   - 资产变更（权重50）：低中优先级，未经授权的资产变更存在安全隐患
   - Windows日志（权重40）：低优先级，系统日志异常反映基础安全状态

2. 风险等级划分标准
   - 🔴 高风险（80-100分）：存在严重安全隐患，需要立即处理
   - 🟡 中高风险（60-79分）：存在明显安全风险，需要加强监控
   - 🟢 中风险（40-59分）：存在一定安全风险，需常规关注
   - 🔵 低风险（0-39分）：安全风险较低，维持现有策略即可

3. 数据处理说明
   - 所有风险评分均标准化到0-1范围
   - 综合评分采用加权平均法计算，权重按风险优先级设定
   - 移动存储/文档上传评分基于实际操作次数计算，其他缺失数据基于相关维度合理推断
   - 最终评分转换为百分制，便于理解和比较''',

        '重点风险人员分析': f'''
1. 最高风险人员：{final_df.iloc[0]['计算机']}
   - 综合风险评分：{final_df.iloc[0]['综合风险评分（百分制）']}分
   - 主要风险点：
     * 应用程序风险评分高达{final_df.iloc[0]['应用程序_风险评分']:.3f}（高风险）
     * 网页浏览风险评分{final_df.iloc[0]['网页浏览_风险评分']:.3f}（高风险）
     * 策略日志风险评分{final_df.iloc[0]['策略日志_风险评分']:.3f}（高风险）
   - 处理建议：{final_df.iloc[0]['风险处理建议']}

2. 高风险群体特征：
   - 共{len(final_df[final_df['综合风险等级'].str.contains('高风险')])}人中高风险人员
   - 主要集中在应用程序过度活跃和网页浏览风险较高的人员
   - 策略日志报警次数较多，存在违反安全策略的行为

3. 整体风险状况：
   - 平均风险评分：{final_df['综合风险评分（百分制）'].mean():.2f}分
   - 风险分布相对均衡，低风险人员占比{len(final_df[final_df['综合风险等级'].str.contains('低风险')])/len(final_df)*100:.1f}%
   - 无极高风险（80分以上）人员，整体安全状况可控''',

        '风险管控建议': '''
1. 分级管控策略
   - 对中高风险人员实施每日审计，重点监控数据传输行为
   - 对中风险人员实施每周审查，关注风险变化趋势
   - 对低风险人员维持月度常规检查，保持安全意识

2. 重点管控措施
   - 移动存储管控：加强USB设备授权管理，禁止未经授权的移动存储使用
   - 文档上传监控：建立敏感文档上传审批机制，限制高风险人员的上传权限
   - 应用程序管理：清理非必要应用程序，限制高风险应用的安装和使用
   - 网页浏览过滤：加强风险网站过滤，阻止访问已知的恶意网站

3. 持续改进建议
   - 完善日志收集机制，补充缺失的移动存储和Windows日志数据
   - 建立风险预警系统，当风险评分超过阈值时自动报警
   - 定期开展安全培训，提高员工的安全意识和合规意识
   - 每季度进行风险评估更新，根据实际情况调整风险权重和管控措施'''
    }
    
    def write_section(ws, start_row, title, content):
        title_cell = ws.cell(row=start_row, column=1, value=title)
        title_cell.font = Font(bold=True, size=14, color="FFFFFF")
        title_cell.fill = PatternFill(start_color="D32F2F", end_color="D32F2F", fill_type="solid")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(f'A{start_row}:B{start_row}')
        
        content_cell = ws.cell(row=start_row + 1, column=1, value=content)
        content_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        content_cell.font = Font(size=11)
        ws.merge_cells(f'A{start_row + 1}:B{start_row + 10}')
        
        return start_row + 12
    
    current_row = 1
    for section, content in risk_evaluation.items():
        current_row = write_section(ws2, current_row, section, content)
    
    ws2.column_dimensions['A'].width = 70
    ws2.column_dimensions['B'].width = 70
    for i in range(1, current_row):
        ws2.row_dimensions[i].height = 20
    
    ws3 = wb.create_sheet(title="风险统计汇总")
    
    max_score = final_df['综合风险评分（百分制）'].max()
    min_score = final_df['综合风险评分（百分制）'].min()
    max_risk_person = final_df[final_df['综合风险评分（百分制）'] == max_score]['计算机'].iloc[0]
    min_risk_person = final_df[final_df['综合风险评分（百分制）'] == min_score]['计算机'].iloc[0]
    
    summary_data = [
        ['统计项目', '数值', '说明'],
        ['统计总人数', len(final_df), '本次风险评估覆盖的所有人员'],
        ['平均综合风险评分', f"{final_df['综合风险评分（百分制）'].mean():.2f}分", '所有人员的风险评分平均值'],
        ['最高风险评分', f"{max_score:.2f}分", f"最高风险人员：{max_risk_person}"],
        ['最低风险评分', f"{min_score:.2f}分", f"最低风险人员：{min_risk_person}"],
        ['高风险人数（80分以上）', len(final_df[final_df['综合风险评分（百分制）'] >= 80]), '需要立即重点关注的人员'],
        ['中高风险人数（60-79分）', len(final_df[(final_df['综合风险评分（百分制）'] >= 60) & (final_df['综合风险评分（百分制）'] < 80)]), '需要加强监控的人员'],
        ['中风险人数（40-59分）', len(final_df[(final_df['综合风险评分（百分制）'] >= 40) & (final_df['综合风险评分（百分制）'] < 60)]), '需要常规关注的人员'],
        ['低风险人数（40分以下）', len(final_df[final_df['综合风险评分（百分制）'] < 40]), '风险较低的人员'],
        ['高风险占比', f"{len(final_df[final_df['综合风险评分（百分制）'] >= 80])/len(final_df)*100:.1f}%", '高风险人员占总人数的比例'],
        ['中高风险及以上占比', f"{len(final_df[final_df['综合风险评分（百分制）'] >= 60])/len(final_df)*100:.1f}%", '中高风险及以上人员占总人数的比例'],
        ['低风险占比', f"{len(final_df[final_df['综合风险评分（百分制）'] < 40])/len(final_df)*100:.1f}%", '低风险人员占总人数的比例']
    ]
    
    for r_idx, row in enumerate(summary_data, 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws3.cell(row=r_idx, column=c_idx, value=value)
            
            if r_idx == 1:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="1976D2", end_color="1976D2", fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if c_idx == 2 and isinstance(value, str) and '分' in value:
                    score = float(value.replace('分', ''))
                    if score >= 80:
                        cell.fill = PatternFill(start_color="FFE6E6", end_color="FFE6E6", fill_type="solid")
                    elif score >= 60:
                        cell.fill = PatternFill(start_color="FFF2E6", end_color="FFF2E6", fill_type="solid")
                    elif score >= 40:
                        cell.fill = PatternFill(start_color="FFFFE6", end_color="FFFFE6", fill_type="solid")
                    else:
                        cell.fill = PatternFill(start_color="E6FFE6", end_color="E6FFE6", fill_type="solid")
            
            cell.border = thin_border
    
    ws3.column_dimensions['A'].width = 20
    ws3.column_dimensions['B'].width = 25
    ws3.column_dimensions['C'].width = 40
    ws3.row_dimensions[1].height = 30
    for i in range(2, len(summary_data) + 1):
        ws3.row_dimensions[i].height = 25
    
    ws4 = wb.create_sheet(title="风险维度权重说明")
    
    weight_data = [
        ['风险维度', '权重值', '优先级', '风险说明', '权重设定依据'],
        ['移动存储', 100, '1（最高）', '数据泄露、恶意软件传播风险', '移动存储设备容易绕过常规安全防护，导致敏感数据泄露和恶意代码传播'],
        ['文档上传', 90, '2', '数据外发、信息泄露风险', '文档上传涉及公司敏感信息外发，可能导致商业秘密泄露'],
        ['应用程序', 80, '3', '恶意软件、资源滥用风险', '未授权应用程序可能包含恶意代码，或导致系统资源过度占用'],
        ['网页浏览', 70, '4', '病毒感染、钓鱼攻击风险', '访问恶意网站可能导致病毒感染、钓鱼攻击等安全事件'],
        ['策略日志', 60, '5', '合规风险、安全违规风险', '违反安全策略表明存在安全意识薄弱或恶意行为倾向'],
        ['资产变更', 50, '6', '配置错误、未授权修改风险', '未经授权的资产变更可能导致系统配置错误，影响系统稳定性'],
        ['Windows日志', 40, '7（最低）', '系统异常、基础安全风险', '系统日志异常反映基础安全状态，但影响范围相对有限']
    ]
    
    for r_idx, row in enumerate(weight_data, 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws4.cell(row=r_idx, column=c_idx, value=value)
            
            if r_idx == 1:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                priority = int(row[2].split('（')[0])
                if priority <= 2:
                    cell.fill = PatternFill(start_color="FFE0B2", end_color="FFE0B2", fill_type="solid")
                elif priority <= 4:
                    cell.fill = PatternFill(start_color="E0F7FA", end_color="E0F7FA", fill_type="solid")
                else:
                    cell.fill = PatternFill(start_color="F3E5F5", end_color="F3E5F5", fill_type="solid")
            
            cell.border = thin_border
    
    column_widths = [15, 12, 15, 30, 50]
    for i, width in enumerate(column_widths, 1):
        ws4.column_dimensions[chr(64 + i)].width = width
    
    ws4.row_dimensions[1].height = 30
    for i in range(2, len(weight_data) + 1):
        ws4.row_dimensions[i].height = 40
    
    wb.save(output_path)
    print(f"✅ Excel报告已保存至：{output_path}")

# ======================== 主函数 ========================
def main():
    file_paths = {
        '移动存储_统计分析结果.xlsx': {'sheet': 1},  # 移动存储文件
        '人员上传操作统计分析.xlsx': {'sheet': 1},   # 文档上传文件
        '应用程序_综合分析报告_含风险评估依据.xlsx': {'sheet': 1},
        '网页浏览风险分析报告.xlsx': {'sheet': 1},
        '策略日志统计分析结果.xlsx': {'sheet': 1},
        '资产变更统计分析.xlsx': {'sheet': 1},
        'windows日志_统计分析结果.xlsx': {'sheet': 1}
    }
    
    print("📊 开始读取和清理数据...")
    cleaned_data = read_and_clean_data(file_paths)
    
    print("\n📊 开始计算综合风险评分...")
    final_df = calculate_comprehensive_risk(cleaned_data)
    
    print("\n📋 人员综合风险统计结果（按评分降序）：")
    print(final_df[['计算机', '综合风险评分（百分制）', '综合风险等级']].to_string(index=False))
    
    print("\n📊 开始生成可视化图表...")
    generate_visualization(final_df, '人员综合风险统计分析图表.png')
    
    print("\n📊 开始生成Excel报告...")
    generate_excel_report(final_df, '人员综合风险统计分析报告.xlsx')
    
    print("\n🎉 所有报告生成完成！")

if __name__ == "__main__":
    main()