import requests

def test_bot(token):
    url = f"https://api.telegram.org/bot{token}/getMe"
    res = requests.get(url)
    if res.status_code == 200:
        print("✅ Token đúng! Bot tên là:", res.json()['result']['first_name'])
    else:
        print("❌ Token sai hoặc Bot chưa được kích hoạt.")

# Gọi hàm này trước khi gửi tin để check cho chắc ăn nhé đại ca!
test_bot("8741654866:AAG7RaTINO5U1BRFB9BNo_1bihCxjyUi290")