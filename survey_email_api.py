import os
import requests
import time
import random
import pandas as pd
from dotenv import load_dotenv
import base64
import json


class GraphApiEmailSender:
    def __init__(self):
        """
        初始化邮件发送器，使用Microsoft Graph API
        """
        self.client_id = os.getenv("CLIENT_ID")
        self.client_secret = os.getenv("CLIENT_SECRET")
        self.tenant_id = os.getenv("TENANT_ID")
        self.sender_email = os.getenv("SENDER_EMAIL")
        self.access_token = None
        self.token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        self.graph_url = "https://graph.microsoft.com/v1.0"

    def get_access_token(self):
        """获取访问令牌"""
        try:
            # 准备请求数据
            data = {
                'client_id': self.client_id,
                'client_secret': self.client_secret,
                'scope': 'https://graph.microsoft.com/.default',
                'grant_type': 'client_credentials'
            }

            # 发送请求获取令牌
            response = requests.post(self.token_url, data=data)

            if response.status_code == 200:
                token_data = response.json()
                self.access_token = token_data['access_token']
                print(f"✅ 成功获取访问令牌")
                return True
            else:
                print(f"❌ 获取访问令牌失败: {response.status_code} - {response.text}")
                return False
        except Exception as e:
            print(f"❌ 获取访问令牌异常: {str(e)}")
            return False

    def read_excel_data(self, file_path, sheet_name=0):
        """
        从Excel文件读取邮件信息
        :param file_path: Excel文件路径
        :param sheet_name: 工作表名称或索引
        :return: DataFrame对象
        """
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            # 检查必要的列是否存在
            required_columns = ['评估人姓名', '员工姓名', '收件人邮箱', '评估链接']
            missing_columns = [col for col in required_columns if col not in df.columns]

            if missing_columns:
                raise ValueError(f"Excel文件缺少必要列: {missing_columns}")

            print(f"✅ 成功读取Excel文件，共{len(df)}条记录")
            return df
        except Exception as e:
            print(f"❌ 读取Excel文件失败: {str(e)}")
            return None

    def create_email_content(self, recipient_name, employee_name, assessment_link):
        """
        创建邮件正文内容
        :param recipient_name: 评估人姓名
        :param employee_name: 员工姓名
        :param assessment_link: 评估链接
        :return: 邮件正文内容
        """
        template = f"""
        <html>
        <body>
            <p>尊敬的 <strong>{recipient_name}</strong>，</p>
            <br>
            <p>您好！</p>
            <br>
            <p>为支持员工的持续成长与发展，我们即将开展2025年度的年终360度评估工作。您是 <strong>{employee_name}</strong> 的重要评估人之一，我们诚挚邀请您花几分钟时间为他/她提供宝贵、真实的反馈。</p>
            <br>
            <p>本次评估将围绕公司的文化-合规守正、以人为本、长期共赢、持续创新等多个维度展开。您的反馈将直接帮助 <strong>{employee_name}</strong> 全面了解自身优势与提升空间，制定更有针对性的个人发展计划。</p>
            <br>
            <p><strong>📌 重要说明：</strong></p>
            <p style="text-indent: 2em; margin-left: 2em;"> • <strong>全程匿名</strong>：您的所有反馈将严格保密，报告汇总后仅以匿名形式呈现，<strong>{employee_name}</strong> 无法看到您的具体评价。</p>
            <p style="text-indent: 2em; margin-left: 2em;"> • <strong>真实坦诚</strong>：我们鼓励您基于事实与观察，提供具体、建设性的意见——这不仅是对同事的负责，更是对公司人才发展的支持。</p>
            <p style="text-indent: 2em; margin-left: 2em;"> • <strong>截止时间</strong>：请于2025年12月31日（星期三）前完成评估。</p>
            <p><a href="{assessment_link}" style="color: #1155CC; text-decoration: underline; display: inline-block;">🔗 点击此处立即填写评估表</a></p>
            <p>您的参与对 <strong>{employee_name}</strong> 的成长至关重要。如有任何疑问，请随时联系HR团队。</p>
            <br>
            <p>感谢您拨冗支持！期待您的真诚反馈。</p>
        </body>
        </html>
        """

        # 填充模板
        content = template.format(
            recipient_name=recipient_name,
            employee_name=employee_name,
            assessment_link=assessment_link
        )
        return content

    def send_single_email(self, recipient_email, recipient_name, employee_name, assessment_link,
                          subject="2025年度年终360度评估邀请"):
        """
        通过Microsoft Graph API发送单封邮件
        :param recipient_email: 收件人邮箱
        :param recipient_name: 评估人姓名
        :param employee_name: 员工姓名
        :param assessment_link: 评估链接
        :param subject: 邮件主题
        :return: 发送结果
        """
        try:
            # 构建邮件内容
            email_content = self.create_email_content(recipient_name, employee_name, assessment_link)

            # 构建请求数据
            message_data = {
                "message": {
                    "subject": subject,
                    "body": {
                        "contentType": "HTML",
                        "content": email_content
                    },
                    "toRecipients": [
                        {
                            "emailAddress": {
                                "address": recipient_email,
                                "name": f"{recipient_name}"
                            }
                        }
                    ],
                    "from": {
                        "emailAddress": {
                            "address": self.sender_email,
                            "name": "HR团队"
                        }
                    }
                },
                "saveToSentItems": True
            }

            # 设置请求头
            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Content-Type': 'application/json'
            }

            # 发送邮件
            url = f"{self.graph_url}/users/{self.sender_email}/sendMail"
            response = requests.post(url, headers=headers, json=message_data)

            if response.status_code == 202:  # 202 Accepted 表示邮件已接受发送
                print(f"✅ 邮件已发送至: {recipient_email} ({recipient_name}) - {employee_name}")
                return True
            else:
                print(f"❌ 发送邮件失败 - {recipient_email}: {response.status_code} - {response.text}")
                return False

        except Exception as e:
            print(f"❌ 发送邮件异常 - {recipient_email}: {str(e)}")
            return False

    def send_bulk_emails(self, excel_file, subject="2025年度年终360度评估邀请", delay_range=(2, 4)):
        """
        批量发送邮件
        :param excel_file: Excel文件路径
        :param subject: 邮件主题
        :param delay_range: 发送间隔时间范围（秒）
        """
        # 读取Excel数据
        df = self.read_excel_data(excel_file)
        if df is None:
            return

        # 获取访问令牌
        if not self.get_access_token():
            print("❌ 无法获取访问令牌，请检查配置")
            print("💡 检查项:")
            print("   1. CLIENT_ID, CLIENT_SECRET, TENANT_ID 是否正确配置")
            print("   2. 应用是否已注册并配置了Mail.Send权限")
            print("   3. 应用是否已获得管理员同意")
            return

        success_count = 0
        fail_count = 0

        try:
            for index, row in df.iterrows():
                recipient_name = str(row.get('评估人姓名', '评估人')).strip()
                employee_name = str(row.get('员工姓名', '员工')).strip()
                recipient_email = str(row.get('收件人邮箱', '')).strip()
                assessment_link = str(row.get('评估链接', '')).strip()

                # 验证必要字段
                if not recipient_email or not assessment_link or not recipient_name or not employee_name:
                    print(f"⚠️  第{index + 1}行数据不完整，跳过发送 - 评估人: {recipient_name}, 员工: {employee_name}")
                    continue

                # 发送邮件
                if self.send_single_email(recipient_email, recipient_name, employee_name, assessment_link, subject):
                    success_count += 1
                else:
                    fail_count += 1

                # 添加随机延迟，避免API调用频率限制
                delay = random.uniform(delay_range[0], delay_range[1])
                time.sleep(delay)

        finally:
            print("✅ 批量发送完成")

        print(f"\n邮件发送完成！")
        print(f"✅ 成功发送: {success_count} 封")
        print(f"❌ 发送失败: {fail_count} 封")


def main():
    # 示例使用
    print("360度评估邮件批量发送工具 (Microsoft Graph API版本)")
    print("=" * 70)

    # 创建邮件发送器实例
    sender = GraphApiEmailSender()

    # 指定Excel文件路径
    excel_file = os.getcwd() + '\\file.xlsx'

    # 设置邮件主题
    subject = "2025年度年终360度评估邀请"

    # 开始批量发送
    sender.send_bulk_emails(excel_file, subject)


if __name__ == "__main__":
    # 提供使用说明
    # print("使用说明 (Microsoft Graph API版本):")
    # print("1. Excel文件应包含以下列：'评估人姓名', '员工姓名', '收件人邮箱', '评估链接'")
    # print("2. 需要配置以下环境变量:")
    # print("   - CLIENT_ID: Azure应用注册的应用(客户端)ID")
    # print("   - CLIENT_SECRET: Azure应用注册的客户端密钥")
    # print("   - TENANT_ID: Azure租户ID")
    # print("   - SENDER_EMAIL: 发件人邮箱地址")
    # print("3. Azure应用需要配置Mail.Send权限并获得管理员同意")
    # print("4. 通过Graph API发送，绕过SMTP限制")
    # print()

    load_dotenv()
    main()