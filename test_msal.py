# שמור כ test_msal.py
import msal

#CLIENT_ID = "7bc0e47e-1c3b-4819-8965-d73956275975"
CLIENT_ID = "9621aecd-afdf-41ee-9989-0a557209c32a"
AUTHORITY = "https://login.microsoftonline.com/consumers"  # או common/tenantId
SCOPES = ["User.Read", "Mail.Read", "email"]

app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
flow = app.initiate_device_flow(scopes=SCOPES)
print(flow)  # אמור להכיל user_code
