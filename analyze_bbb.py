import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

def generate_report(file_path, output_file):
    # Load data
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error reading file: {e}")
        return

    # Data Preprocessing
    df['日期'] = pd.to_datetime(df['日期'])
    df['Month'] = df['日期'].dt.to_period('M').astype(str)
    
    # Initialize HTML content
    html_content = f"""
    <html>
    <head>
        <title>公文数据分析报告</title>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 20px; background-color: #f4f4f9; }}
            h1 {{ text-align: center; color: #333; }}
            .summary {{ display: flex; justify-content: space-around; background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); margin-bottom: 20px; }}
            .metric {{ text-align: center; }}
            .metric h2 {{ margin: 0; color: #007bff; }}
            .metric p {{ margin: 5px 0 0; color: #666; }}
            .chart-container {{ background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); margin-bottom: 20px; }}
        </style>
    </head>
    <body>
        <h1>公文数据可视化分析报告</h1>
        
        <div class="summary">
            <div class="metric">
                <h2>{len(df)}</h2>
                <p>总文件数</p>
            </div>
            <div class="metric">
                <h2>{df['机构'].nunique()}</h2>
                <p>涉及机构数</p>
            </div>
            <div class="metric">
                <h2>{df['签发人'].nunique()}</h2>
                <p>签发人数</p>
            </div>
            <div class="metric">
                <h2>{df['拟稿人'].nunique()}</h2>
                <p>拟稿人数</p>
            </div>
            <div class="metric">
                <h2>{df['日期'].min().date()} - {df['日期'].max().date()}</h2>
                <p>时间跨度</p>
            </div>
        </div>
    """

    # 1. Monthly Trend
    monthly_counts = df.groupby('Month').size().reset_index(name='Count')
    fig_trend = px.line(monthly_counts, x='Month', y='Count', title='每月发文数量趋势', markers=True)
    html_content += f'<div class="chart-container">{fig_trend.to_html(full_html=False, include_plotlyjs="cdn")}</div>'

    # 2. Top Issuers (签发人)
    top_issuers = df['签发人'].value_counts().nlargest(10).reset_index()
    top_issuers.columns = ['签发人', '文件数']
    fig_issuer = px.bar(top_issuers, x='签发人', y='文件数', title='前10位签发人发文量', color='文件数')
    html_content += f'<div class="chart-container">{fig_issuer.to_html(full_html=False, include_plotlyjs=False)}</div>'

    # 3. Department Distribution (机构)
    dept_counts = df['机构'].value_counts().reset_index()
    dept_counts.columns = ['机构', '文件数']
    fig_dept = px.pie(dept_counts, names='机构', values='文件数', title='各机构发文占比', hole=0.4)
    html_content += f'<div class="chart-container">{fig_dept.to_html(full_html=False, include_plotlyjs=False)}</div>'

    # 4. Retention Period (保管期限) & Submission Unit (报送单位)
    # Using subplots for these two
    fig_sub = make_subplots(rows=1, cols=2, specs=[[{'type': 'domain'}, {'type': 'xy'}]], 
                            subplot_titles=['保管期限分布', '前10位报送单位'])
    
    # Retention Pie
    retention_counts = df['保管期限'].value_counts().reset_index()
    retention_counts.columns = ['保管期限', 'Count']
    fig_sub.add_trace(go.Pie(labels=retention_counts['保管期限'], values=retention_counts['Count'], name="保管期限"), 1, 1)

    # Submission Unit Bar
    sub_unit_counts = df['报送单位'].value_counts().nlargest(10).reset_index()
    sub_unit_counts.columns = ['报送单位', 'Count']
    fig_sub.add_trace(go.Bar(x=sub_unit_counts['报送单位'], y=sub_unit_counts['Count'], name="报送单位"), 1, 2)
    
    fig_sub.update_layout(title_text="保管期限与报送单位分析")
    html_content += f'<div class="chart-container">{fig_sub.to_html(full_html=False, include_plotlyjs=False)}</div>'

    # 5. Top Drafters (拟稿人)
    top_drafters = df['拟稿人'].value_counts().nlargest(15).reset_index()
    top_drafters.columns = ['拟稿人', '文件数']
    fig_drafter = px.bar(top_drafters, x='拟稿人', y='文件数', title='前15位拟稿人工作量', color='文件数')
    html_content += f'<div class="chart-container">{fig_drafter.to_html(full_html=False, include_plotlyjs=False)}</div>'

    html_content += """
    </body>
    </html>
    """

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"Report generated successfully: {output_file}")

if __name__ == "__main__":
    generate_report('bbb.xlsx', 'bbb_analysis_report.html')
