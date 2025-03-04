# 发票OCR处理工具

本项目使用腾讯云OCR服务对PDF格式的发票进行识别，并提取关键信息。识别出的数据自动保存为Excel文件，并对PDF文件进行重命名，以便于管理和查询。

## 功能介绍
- **PDF转换**: 将PDF发票转换为图片以便OCR识别
- **OCR识别**: 通过腾讯云OCR接口提取发票信息
- **文件重命名**: 按照发票信息（姓名、销售方、发票号、金额）重新命名PDF文件
- **数据存储**: 识别结果保存至Excel文件，方便后续分析

## 环境依赖
请确保已安装以下Python库：
```bash
pip install json os pandas pdf2image tencentcloud pillow openpyxl
