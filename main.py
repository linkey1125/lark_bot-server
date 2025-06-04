import re
import openpyxl

# メール内容の解析
def parse_email(text):
    # 正規表現を使って各項目を抽出
    project_name = re.search(r'【案件名】\s*(.*)', text)
    work_content = re.search(r'【作業内容】\s*(.*)', text)
    required_skills = re.search(r'【必須スキル】\s*([\s\S]*?)\n【', text)
    num_recruited = re.search(r'【募集人数】\s*(.*)', text)
    period = re.search(r'【期間】\s*(.*)', text)
    location = re.search(r'【勤務地】\s*(.*)', text)
    other_info = re.search(r'【勤務時間】\s*([\s\S]*)', text)

    # 抽出した情報を辞書として返す
    return {
        '案件名': project_name.group(1).strip() if project_name else '未記入',
        '作業内容': work_content.group(1).strip() if work_content else '未記入',
        '募集要件': required_skills.group(1).strip() if required_skills else '未記入',
        '募集人数': num_recruited.group(1).strip() if num_recruited else '未記入',
        '期間': period.group(1).strip() if period else '未記入',
        '勤務場所': location.group(1).strip() if location else '未記入',
        'その他': other_info.group(1).strip() if other_info else '未記入',
    }

# メール内容を入力
email_text_1 = """
①
【案件名】
某業界向けSAP UI5開発
【期間】
7月～9月末予定（延長、または他案件シフトあり）
【募集人数】
 １名
【勤務地】
虎ノ門
【作業内容】
UI5での画面新規作成

【必須スキル】
・UI5での開発経験
・日本語での円滑なコミュニケーション

【勤務時間】
現場の勤務時間に従い、
現場常駐（在宅は要相談）
"""

# メールを解析してデータを表示
parsed_data = parse_email(email_text_1)
print(parsed_data)

# Excelファイルに書き出す
def export_to_excel(data, filename="output.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['案件名', '作業内容', '募集要件', '募集人数', '期間', '勤務場所', 'その他'])
    ws.append([
        data['案件名'],
        data['作業内容'],
        data['募集要件'],
        data['募集人数'],
        data['期間'],
        data['勤務場所'],
        data['その他']
    ])
    wb.save(filename)

# 解析したデータをExcelに出力
export_to_excel(parsed_data)
