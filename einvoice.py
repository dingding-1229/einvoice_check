import json
import os
import pandas as pd
from pdf2image import convert_from_path
from tencentcloud.common import credential
from tencentcloud.common.profile.client_profile import ClientProfile
from tencentcloud.common.profile.http_profile import HttpProfile
from tencentcloud.ocr.v20181119 import ocr_client, models
from tencentcloud.common.exception.tencent_cloud_sdk_exception import TencentCloudSDKException
import base64
import re
from PIL import Image

# 规范化文件名（去除特殊字符）
def sanitize_filename(name):
    return re.sub(r'[\/:*?"<>|]', '_', name)

# 将图像转换为 Base64
def image_to_base64(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode('utf-8')

# 腾讯云凭证
cred = credential.Credential("Your Tencent ID", "Your Tencent Key")
httpProfile = HttpProfile()
httpProfile.endpoint = "ocr.tencentcloudapi.com"
clientProfile = ClientProfile()
clientProfile.httpProfile = httpProfile
client = ocr_client.OcrClient(cred, "", clientProfile)

# 设定文件夹路径，包含多个PDF文件
pdf_folder = "发票"

# 你的姓名
user_name = "丁逸聪"


# 存储所有 OCR 识别结果
all_results = []


# 遍历 PDF 文件
for pdf_filename in os.listdir(pdf_folder):
    if pdf_filename.endswith(".pdf"):
        pdf_path = os.path.join(pdf_folder, pdf_filename)
        print(f"Processing {pdf_filename}...")

        # 将 PDF 转换为图像
        images = convert_from_path(pdf_path)

        # 逐页处理
        for i, image in enumerate(images):
            image_path = f"{pdf_filename}_page_{i + 1}.png"
            image.save(image_path, "PNG")

            try:
                # OCR 识别
                req = models.RecognizeGeneralInvoiceRequest()
                params = {"ImageBase64": image_to_base64(image_path)}
                req.from_json_string(json.dumps(params))
                resp = client.RecognizeGeneralInvoice(req)
                result_json = json.loads(resp.to_json_string())

                # 获取发票信息
                invoice_info = result_json["MixedInvoiceItems"][0]["SingleInvoiceInfos"]
                vat_special = invoice_info.get("VatElectronicSpecialInvoiceFull")  # 专用发票
                vat_common = invoice_info.get("VatElectronicInvoiceFull")  # 普通发票
                invoice = vat_special or vat_common  # 选取不为空的发票

                if invoice:
                    # 提取发票信息
                    seller = invoice.get("Seller", "未知销售方")
                    invoice_number = invoice.get("Number", "未知发票号")
                    total_amount = invoice.get("Total", "未知金额")

                    # 规范化
                    seller_clean = sanitize_filename(seller)
                    invoice_number_clean = sanitize_filename(invoice_number)
                    total_amount_clean = sanitize_filename(total_amount)

                    # 生成新文件名
                    new_pdf_name = f"{user_name}-{seller_clean}-{invoice_number_clean}-{total_amount_clean}.pdf"
                    new_pdf_path = os.path.join(pdf_folder, new_pdf_name)

                    # 重命名 PDF 文件
                    os.rename(pdf_path, new_pdf_path)
                    # print(f"Renamed: {pdf_filename} -> {new_pdf_name}")

                    # 记录信息
                    all_results.append({
                        "OriginalFile": pdf_filename,
                        "NewFileName": new_pdf_name,
                        "InvoiceType": invoice.get("Title", ""),
                        "InvoiceNumber": invoice_number,
                        "Date": invoice.get("Date", ""),
                        "Seller": seller,
                        "SellerTaxID": invoice.get("SellerTaxID", ""),
                        "Buyer": invoice.get("Buyer", ""),
                        "BuyerTaxID": invoice.get("BuyerTaxID", ""),
                        "PretaxAmount": invoice.get("PretaxAmount", ""),
                        "Tax": invoice.get("Tax", ""),
                        "TotalAmount": total_amount,
                        "Issuer": invoice.get("Issuer", ""),
                        "Remark": invoice.get("Remark", ""),
                        "TaxRate": invoice.get("VatElectronicItems", [{}])[0].get("TaxRate", ""),
                    })

                else:
                    print(f"No invoice found on page {i + 1} of {pdf_filename}")

            except TencentCloudSDKException as err:
                print(f"Error on page {i + 1} of {pdf_filename}: {err}")

            # 清理图片
            os.remove(image_path)

# 结果保存 Excel
df = pd.DataFrame(all_results)
output_path = "ocr_results.xlsx"
df.to_excel(output_path, index=False)

print(f"OCR results saved to {output_path}")
