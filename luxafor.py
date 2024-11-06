import requests

class Luxafor:
    @staticmethod
    def set_color(color, luxafor_id):
        url = "https://api.luxafor.com/webhook/v1/actions/solid_color"
        headers = {
            "Content-Type": "application/json"
        }
        payload = {
            "userId": luxafor_id,
            "actionFields": {
                "color": color
            }
        }
        response = requests.post(url, json=payload, headers=headers)

        if response.status_code == 200:
            print(f"Luxafor color set to {color} successfully.")
        else:
            print(f"Failed to set Luxafor color to {color}. Error: {response.text}")