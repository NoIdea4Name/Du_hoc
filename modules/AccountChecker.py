import requests

class AccountChecker:
    def run(self, sfid: str):
        sfid = sfid.replace("-", "")
        r = requests.get(f"https://ss2.sfcollege.edu/sr/account/verification/{sfid}/status")
        print(r.text, r.status_code)
        try:
            if r.json()['payload']['accountStatus'] == "ACTIVE":
                return True
            else:
                return False
                
        except Exception as e:
            print("ERROR:", e)
            return None