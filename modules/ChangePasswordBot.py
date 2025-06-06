import re
import time
import requests
from modules.TempMailClient import TempMailClient

FORGOT_PW_URL = "https://ss2.sfcollege.edu/pwmanager/api/forgotpw"
AUTHCODE_URL  = "https://ss2.sfcollege.edu/pwmanager/api/authcode"
CHANGEPW_URL  = "https://ss2.sfcollege.edu/pwmanager/api/changepw"


class ChangePasswordBot():
    def __init__(self, token, wait_sec=10):
        self.token = token
        self.client = TempMailClient(token)
        self.wait_sec = wait_sec

    def extract_code(self, html):
        """Lấy mã authcode từ HTML."""
        m = re.search(r"<b>(.*?)</b>", html)
        return m.group(1) if m else None
    
    def run(self, email, sfid, password, wait_sec_get_mail=10):
        session = requests.Session()
        sfid = sfid.replace("-", "") if "-" in sfid else sfid

        resp = session.post(FORGOT_PW_URL, json={"sfid": sfid, "email": email}).json()
        print(resp)
        if not resp.get("success"):
            return None

        user_name = email.split("@")[0]
        domain = email.split("@")[1]
        mail_id = self.client.create_temp_email(user_name, domain).get("id")
    
        if not mail_id:
            return None

        time.sleep(wait_sec_get_mail)
        msg_id = self.client.get_message_by_match(mail_id, 
                                            by_sender="no-reply@sfcollege.edu", 
                                            by_subject="eSantaFe Password Change Auth Code")
        
        if not msg_id:
            return None

        msg = self.client.read_message(msg_id)
        if not msg:
            return None
            
        code = self.extract_code(msg.get("body"))
        if not code:
            return None
        
        resp = session.post(AUTHCODE_URL, json={"authcode": code, "sfid": sfid}).json()
        print(resp)
        if not resp.get("success"):
            return None

        resp = session.post(CHANGEPW_URL,
                        json={"authcode": code, "sfid": sfid, "pw": password}).json()  
        print(resp)
        return password if resp.get("success") else None
        
    def generate_password(self):
        resp = requests.get("https://www.dinopass.com/password/strong")
        if resp.status_code != 200:
            return None
        return resp.text.strip()