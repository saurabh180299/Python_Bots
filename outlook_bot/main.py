from fastapi import FastAPI
import win32com.client
from fastapi.middleware.cors import CORSMiddleware
import pythoncom
import requests
from fastapi.responses import JSONResponse
# import json

app = FastAPI()

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=['*'],  # Replace with the appropriate origins in production
    allow_methods=['GET', 'POST', 'PUT', 'DELETE'],
    allow_headers=['*'],
)

@app.get("/emails")
def get_recent_emails():
    pythoncom.CoInitialize()  # Initialize the COM library
    
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 6 represents the index of the inbox folder

    recent_emails = []
    counter = 0
    for email in reversed(inbox.Items):
        subject = email.Subject
        sender = email.SenderName
        received_time = email.ReceivedTime
        content = email.body
        
        recent_emails.append({
            "subject": subject,
            "sender": sender,
            "received_time": received_time,
            "content": content
        })

        counter += 1
        if counter >= 5:
            break

    pythoncom.CoUninitialize()  # Uninitialize the COM library
    
    return recent_emails


@app.get("/search_stack_overflow")
async def search_stack_overflow(query: str):
    url = f"https://api.stackexchange.com/2.3/search?order=desc&sort=activity&site=stackoverflow&intitle={query}"
    response = requests.get(url)

    if response.status_code == 200:
        data = response.json()
        if data["items"]:
            links = [item["link"] for item in data["items"]]
            return {"links": links}
        else:
            return {"message": "No answers on Stack Overflow."}
    else:
        return JSONResponse(
            status_code=response.status_code,
            content={"message": "Error searching on Stack Overflow."},
        )

