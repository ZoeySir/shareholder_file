"""
股东协议分析处理程序
"""
import logging
import os
import re
import sys
import time
from pathlib import Path
import yaml
import json
import pandas as pd
from docx import Document
import requests
from openpyxl import load_workbook

# 日志配置
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename=Path(__file__).parent / 'shareholder_log.log',
    filemode='a',
    encoding='utf-8'
)
logger = logging.getLogger(__name__)

# 配置文件结构示例
"""
config:
  word_paths:  # 支持多个路径
    - /Users/macuser/Documents/contracts/doc1.docx
    - C:\\Contracts\\doc2.docx
  excel_path: /path/to/config.xlsx
  api:
    url: http://api.example.com/analyze
    auth: Bearer xxxx
"""


def load_config(config_path='config.yaml'):
    """加载配置文件"""
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)

        # 获取嵌套的配置项
        app_config = config['config']  # 关键修正点：访问嵌套层级

        # 路径兼容性处理（现在操作正确层级的配置项）
        app_config['excel_path'] = Path(app_config['excel_path']).resolve()
        app_config['word_paths'] = [Path(p).resolve() for p in app_config['word_paths']]

        logger.info("配置文件加载成功")
        return app_config  # 返回已处理的配置字典
    except KeyError as e:
        logger.error(f"配置项缺失: {str(e)}")
        sys.exit(1)
    except Exception as e:
        logger.error(f"配置文件加载失败: {str(e)}")
        sys.exit(1)

def read_word_doc(doc_path):
    """读取Word文档并清理格式"""
    try:
        doc = Document(doc_path)
        full_text = []

        # 清理规则
        page_num_re = re.compile(r'第[一二三四五六七八九十]+条')
        dots_re = re.compile(r'\.{4,}')

        for para in doc.paragraphs:
            text = para.text.strip()

            # 跳过页码行和目录点
            if page_num_re.match(text) or dots_re.search(text):
                continue

            # 合并多余空行
            if text == '' and full_text and full_text[-1] == '\n':
                continue

            full_text.append(text.replace('\t', ' ') + '\n')

        return ''.join(full_text).replace('\n\n', '\n')
    except Exception as e:
        logger.error(f"读取Word文档失败[{doc_path}]: {str(e)}")
        return None


def process_excel_row(row, file_content, api_url, api_auth):
    """处理单行Excel数据"""
    try:
        # 构建Query参数
        query_data = {
            "条款类别": row['条款类别'],
            "条款/内容": row['条款/内容'],
            "填写说明": row['填写说明']
        }
        query_json = json.dumps(query_data, ensure_ascii=False, indent=2)

        # API请求
        headers = {
            'Authorization': api_auth,
            'Content-Type': 'application/json'
        }
        payload = {
            "response_mode": "blocking",
            "user": "haha-shareholder",
            "inputs": {
                "Query": query_json,  # 注意这里保持字符串形式
                "File": file_content
            }
        }

        response = requests.post(
            api_url,
            json=payload,
            headers=headers,
            timeout=600  # 10分钟超时
        )
        response.raise_for_status()

        # 解析响应
        result_json = response.json()
        # for line in response.iter_lines():
        #     if b'"event": "workflow_finished"' in line:
        #         result_data = json.loads(line.decode().split('data: ')[1])
        #         answer = result_data['data']['outputs']['Answer']
        #
        #         # 提取json内容
        #         json_match = re.search(r'```json\n(.*?)\n```', answer, re.DOTALL)
        #         if json_match:
        #             result_json = json.loads(json_match.group(1))
        #         break
        answer = result_json['data']['outputs']['Answer']
        json_match = re.search(r'```json(.*?)```', answer, re.DOTALL)
        if not json_match:
            raise ValueError("未找到有效JSON数据")

        result = json.loads(json_match.group(1).strip())
        return {
            '概括总结': result.get('概括总结', '未提取'),
            '条款编号': result.get('条款编号', '未提取'),
            '条款原文': result.get('条款原文', '未提取'),
            'error': ''
        }
    except Exception as e:
        logger.error(f"处理行数据失败: {str(e)}")
        return {
            '概括总结': '处理失败',
            '条款编号': '处理失败',
            '条款原文': '处理失败',
            'error': str(e)
        }


def main():
    # 加载配置
    config = load_config()

    # 处理所有Word文档
    for word_path in config['word_paths']:
        if not word_path.exists():
            logger.error(f"Word文件不存在: {word_path}")
            continue

        # 读取Word内容
        file_content = read_word_doc(word_path)
        if not file_content:
            continue

        # 处理Excel
        try:
            df = pd.read_excel(config['excel_path'], engine='openpyxl')

            # 创建结果DataFrame
            result_df = df.copy()
            result_df[['概括总结', '条款编号', '条款原文', 'error']] = ''

            # 逐行处理
            for idx, row in df.iterrows():
                logger.info(f"正在处理第{idx + 1}行: {row['条款/内容']}")
                result = process_excel_row(
                    row,
                    file_content,
                    config['api']['url'],
                    config['api']['auth']
                )

                # 更新结果
                for col in ['概括总结', '条款编号', '条款原文', 'error']:
                    result_df.at[idx, col] = result[col]

            # 保存结果
            output_path = word_path.parent / f"{word_path.stem}_result.xlsx"
            result_df.to_excel(output_path, index=False, engine='openpyxl')
            logger.info(f"结果文件已保存: {output_path}")

        except Exception as e:
            logger.error(f"处理Excel失败: {str(e)}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logger.info("用户中断程序")
    except Exception as e:
        logger.error(f"程序异常终止: {str(e)}")
