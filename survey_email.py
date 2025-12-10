import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import pandas as pd
import time
import random


class OutlookEmailSender:
    def __init__(self, sender_email, sender_password):
        """
        初始化邮件发送器
        :param sender_email: 发件人邮箱地址
        :param sender_password: 发件人邮箱密码或应用专用密码
        """
        self.sender_email = sender_email
        self.sender_password = sender_password
        self.smtp_server = "smtp-mail.outlook.com"
        self.smtp_port = 587

    def connect_smtp(self):
        """连接SMTP服务器"""
        try:
            self.server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            self.server.starttls()  # 启用加密传输
            self.server.login(self.sender_email, self.sender_password)
            print(f"成功连接到 {self.sender_email} 的SMTP服务器")
            return True
        except Exception as e:
            print(f"连接SMTP服务器失败: {str(e)}")
            return False

    def disconnect_smtp(self):
        """断开SMTP连接"""
        if hasattr(self, 'server'):
            self.server.quit()
            print("已断开SMTP连接")

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

            print(f"成功读取Excel文件，共{len(df)}条记录")
            return df
        except Exception as e:
            print(f"读取Excel文件失败: {str(e)}")
            return None

    def create_email_content(self, recipient_name, employee_name, assessment_link):
        """
        创建邮件正文内容
        :param recipient_name: 评估人姓名
        :param employee_name: 员工姓名
        :param assessment_link: 评估链接
        :return: 邮件正文内容
        """
        template = """
        <html>
        <head>
            <style>
                body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; }}
                .highlight {{ background-color: #f0f8ff; padding: 10px; border-left: 4px solid #0078d4; }}
                .link-btn {{ 
                    display: inline-block; 
                    padding: 10px 20px; 
                    background-color: #0078d4; 
                    color: white !important; 
                    text-decoration: none; 
                    border-radius: 4px; 
                    margin: 10px 0;
                }}
                .important {{ color: #d32f2f; font-weight: bold; }}
            </style>
        </head>
        <body>
            <p>尊敬的 <strong>{recipient_name}</strong>，</p>
            <br>
            <p>您好！</p>
            <br>
            <p>为支持员工的持续成长与发展，我们即将开展2025年度的年终360度评估工作。您被 <strong>{employee_name}</strong> 指定为重要评估人之一，我们诚挚邀请您花几分钟时间为他/她提供宝贵、真实的反馈。</p>
            <br>
            <p>本次评估将围绕公司的文化-合规守正、以人为本、长期共赢、持续创新等多个维度展开。您的反馈将直接帮助 <strong>{employee_name}</strong> 全面了解自身优势与提升空间，制定更有针对性的个人发展计划。</p>
            <br>
            <div class="highlight">
                <strong>📌 重要说明：</strong><br>
                • <span class="important">全程匿名</span>：您的所有反馈将严格保密，报告汇总后仅以匿名形式呈现，<strong>{employee_name}</strong> 无法看到您的具体评价。<br>
                • <span class="important">真实坦诚</span>：我们鼓励您基于事实与观察，提供具体、建设性的意见——这不仅是对同事的负责，更是对公司人才发展的支持。<br>
                • <span class="important">截止时间</span>：请于2025年12月31日（星期三）前完成评估。
            </div>
            <br>
            <p><a href="{assessment_link}" class="link-btn" target="_blank">🔗 点击此处立即填写评估表</a></p>
            <p style="margin-left: 20px;"><small>{assessment_link}</small></p>
            <br>
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
        发送单封邮件
        :param recipient_email: 收件人邮箱
        :param recipient_name: 评估人姓名
        :param employee_name: 员工姓名
        :param assessment_link: 评估链接
        :param subject: 邮件主题
        :return: 发送结果
        """
        try:
            # 创建邮件对象
            msg = MIMEMultipart('alternative')
            msg['From'] = Header(f"HR团队 <{self.sender_email}>", 'utf-8')
            msg['To'] = Header(f"{recipient_name} <{recipient_email}>", 'utf-8')
            msg['Subject'] = Header(subject, 'utf-8')

            # 创建邮件正文
            body = self.create_email_content(recipient_name, employee_name, assessment_link)
            msg.attach(MIMEText(body, 'html', 'utf-8'))

            # 发送邮件
            self.server.send_message(msg)
            print(f"邮件已发送至: {recipient_email} ({recipient_name}) - {employee_name}")
            return True

        except Exception as e:
            print(f"发送邮件失败 - {recipient_email}: {str(e)}")
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

        # 连接SMTP服务器
        if not self.connect_smtp():
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
                    print(f"第{index + 1}行数据不完整，跳过发送 - 评估人: {recipient_name}, 员工: {employee_name}")
                    continue

                # 发送邮件
                if self.send_single_email(recipient_email, recipient_name, employee_name, assessment_link, subject):
                    success_count += 1
                else:
                    fail_count += 1

                # 添加随机延迟，避免被识别为垃圾邮件
                delay = random.uniform(delay_range[0], delay_range[1])
                time.sleep(delay)

        finally:
            self.disconnect_smtp()

        print(f"\n邮件发送完成！")
        print(f"成功发送: {success_count} 封")
        print(f"发送失败: {fail_count} 封")


def main():
    # 示例使用
    print("360度评估邮件批量发送工具")
    print("=" * 60)

    # 设置发件人信息（请替换为实际邮箱信息）
    SENDER_EMAIL = ""
    SENDER_PASSWORD = ""

    # 创建邮件发送器实例
    sender = OutlookEmailSender(SENDER_EMAIL, SENDER_PASSWORD)

    # 指定Excel文件路径
    excel_file = os.getcwd()  + '\\file.xlsx'

    # 设置邮件主题
    subject = "2025年度年终360度评估邀请"

    # 开始批量发送
    sender.send_bulk_emails(excel_file, subject)


if __name__ == "__main__":
    # 提供使用说明
    # print("使用说明:")
    # print("1. Excel文件应包含以下列：'评估人姓名', '员工姓名', '收件人邮箱', '评估链接'")
    # print("2. 请确保发件人邮箱已开启SMTP服务")
    # print("3. Outlook邮箱需要使用应用专用密码")
    # print("4. 为了防止被识别为垃圾邮件，程序会在每封邮件之间添加随机延迟")
    # print("5. 邮件模板严格按照要求设计，包含所有指定内容")
    # print()

    main()